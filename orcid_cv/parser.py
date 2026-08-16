import os
import re
import json
import logging
import requests
import xmltodict
from typing import Dict, List, Any, Callable
from urllib.parse import urlparse
from collections import defaultdict

from orcid_cv.utils import get_recursive_key, dict_to_list

logger = logging.getLogger("orcid_cv")


def load_xml(xml_path: str) -> Dict[str, Any]:
    """Loads an XML file and converts it into a Python dictionary."""
    with open(xml_path, encoding="utf-8") as xd:
        xml_dict = xmltodict.parse(xd.read())

    # Remove top-level wrapper if it exists
    if len(xml_dict.keys()) == 1:
        return list(xml_dict.values())[0]
    else:
        raise ValueError("XML dictionary has more than one top-level key")


def list_works(orcid_dir: str) -> None:
    """Lists the titles and put-codes of all works in the works directory."""
    works_path = os.path.join(orcid_dir, "works")
    if not os.path.exists(works_path):
        print(f"No works directory found at {works_path}")
        return
        
    work_xml_list = os.listdir(works_path)
    for i, w in enumerate(work_xml_list):
        xml_dict = load_xml(os.path.join(works_path, w))
        title = get_recursive_key(xml_dict, "work:work", "work:title", "common:title")
        put_code = get_recursive_key(xml_dict, "work:work", "@put-code")
        print(f"{i}: {title} ({put_code})")


def load_affiliation(affiliation_path: str) -> Dict[str, Any]:
    """
    Loads a single affiliation record. Employments, educations and services all
    use the same `common:` schema, so one loader covers all three folders.
    """
    affiliation_xml = load_xml(affiliation_path)
    affiliation_dict = {
        "organization": get_recursive_key(
            affiliation_xml, "common:organization", "common:name"
        ),
        "department": get_recursive_key(affiliation_xml, "common:department-name"),
        "role": get_recursive_key(affiliation_xml, "common:role-title"),
        "start_date": get_recursive_key(
            affiliation_xml, "common:start-date", "common:year"
        ),
        "end_date": get_recursive_key(
            affiliation_xml, "common:end-date", "common:year"
        ),
    }
    
    if affiliation_dict["end_date"] == "":
        affiliation_dict["date_range"] = affiliation_dict["start_date"] + " - present"
    elif affiliation_dict["end_date"] == affiliation_dict["start_date"]:
        # Single-year entries (common for service) read better without a range
        affiliation_dict["date_range"] = affiliation_dict["start_date"]
    else:
        affiliation_dict["date_range"] = (
            affiliation_dict["start_date"] + " - " + affiliation_dict["end_date"]
        )

    return affiliation_dict


def load_work(work_path: str) -> Dict[str, Any]:
    """Loads a single work record, extracting metadata, identifiers, and authors."""
    in_work_dict = load_xml(work_path)
    out_work_dict = {
        "type": in_work_dict.get("work:type", ""),
        "title": get_recursive_key(in_work_dict, "work:title", "common:title"),
        "subtitle": "",
        "journal": get_recursive_key(in_work_dict, "work:journal-title"),
        "doi": get_recursive_key(in_work_dict, "common:url"),
        "year": get_recursive_key(
            in_work_dict, "common:publication-date", "common:year"
        ),
        "month": get_recursive_key(
            in_work_dict, "common:publication-date", "common:month"
        ),
        "authors": get_recursive_key(
            in_work_dict, "work:contributors", "work:contributor"
        ),
        "external_ids": []
    }

    # Extract external IDs (specifically DOIs)
    external_ids = get_recursive_key(
        in_work_dict, "common:external-ids", "common:external-id"
    )
    if isinstance(external_ids, list):
        out_work_dict["external_ids"] = [
            d["common:external-id-value"]
            for d in external_ids
            if d.get("common:external-id-type") == "doi"
        ]
    elif isinstance(external_ids, dict):
        out_work_dict["external_ids"] = [external_ids.get("common:external-id-value", "")]

    # Remove empty external IDs
    out_work_dict["external_ids"] = [eid for eid in out_work_dict["external_ids"] if eid]

    # Defaults for sorting
    if out_work_dict["year"] == "":
        out_work_dict["year"] = 0
    if out_work_dict["month"] == "":
        out_work_dict["month"] = 0

    # Extract authors list
    authors = out_work_dict["authors"]
    if authors != "":
        if isinstance(authors, dict):
            out_work_dict["authors"] = [authors.get("work:credit-name")]
        elif isinstance(authors, list):
            out_work_dict["authors"] = [
                item.get("work:credit-name") for item in authors if item
            ]
        out_work_dict["authors"] = [a for a in out_work_dict["authors"] if a]
    else:
        out_work_dict["authors"] = []

    # Random shuffling of keys
    if out_work_dict["type"] in ["software", "conference-presentation"]:
        out_work_dict["subtitle"] = get_recursive_key(
            in_work_dict, "work:title", "common:subtitle"
        )
    
    # Remove author from presentations
    if out_work_dict["type"] in ["public-speech", "conference-presentation"]:
        out_work_dict["authors"] = ""
        
    return out_work_dict


def normalize_title(title: str) -> str:
    """Helper to normalize paper titles for robust matching."""
    if not title:
        return ""
    # Lowercase and remove all non-alphanumeric characters, ignoring whitespace differences
    cleaned = "".join(c.lower() for c in title if c.isalnum() or c.isspace())
    return " ".join(cleaned.split())


def check_duplicates(input_dict: Dict[str, Any]) -> bool:
    """
    Checks if there are duplicate titles or external IDs across preprints and articles.
    """
    eids = []
    titles = []
    for di in input_dict.values():
        if di.get("type") not in ["preprint", "journal-article"]:
            continue
        eids.extend(di.get("external_ids", []))
        titles.append(normalize_title(di.get("title", "")))

    if len(eids) != len(set(eids)) or len(titles) != len(set(titles)):
        return True
    return False


def prune_duplicate_works(work_dict: Dict[str, Any]) -> Dict[str, Any]:
    """
    Groups duplicate preprints and articles using a connected components graph algorithm,
    merging their IDs and choosing the best representative in a single pass.
    """
    # 1. Build adjacency list of connected papers
    adj = defaultdict(list)
    title_to_keys = defaultdict(list)
    eid_to_keys = defaultdict(list)
    
    keys_to_process = []
    for key, w in work_dict.items():
        if w.get("type") not in ["preprint", "journal-article"]:
            continue
        keys_to_process.append(key)
        
        nt = normalize_title(w.get("title", ""))
        if nt:
            title_to_keys[nt].append(key)
        for eid in w.get("external_ids", []):
            if eid:
                eid_to_keys[eid].append(key)
                
    # Connect keys sharing the same normalized title
    for keys in title_to_keys.values():
        for i in range(len(keys) - 1):
            adj[keys[i]].append(keys[i+1])
            adj[keys[i+1]].append(keys[i])
            
    # Connect keys sharing the same external ID
    for keys in eid_to_keys.values():
        for i in range(len(keys) - 1):
            adj[keys[i]].append(keys[i+1])
            adj[keys[i+1]].append(keys[i])
            
    # 2. Find connected components (duplicate groups)
    visited = set()
    components = []
    for key in keys_to_process:
        if key not in visited:
            comp = []
            stack = [key]
            while stack:
                node = stack.pop()
                if node not in visited:
                    visited.add(node)
                    comp.append(node)
                    if node in adj:
                        stack.extend(adj[node])
            components.append(comp)
            
    # 3. For each group with duplicates, merge them into the best representative
    for comp in components:
        if len(comp) <= 1:
            continue
            
        # Priority logic to select the best representative (keep_key)
        # Prefer journal-article, newer year, newer month, longest ID, longest DOI
        def get_priority(k: str):
            w = work_dict[k]
            type_pref = 1 if w.get("type") == "journal-article" else 0
            try:
                year = int(w.get("year", 0))
            except (ValueError, TypeError):
                year = 0
            try:
                month = int(w.get("month", 0))
            except (ValueError, TypeError):
                month = 0
            eids = w.get("external_ids", [])
            max_eid_len = max(len(eid) for eid in eids) if eids else 0
            doi_len = len(w.get("doi", ""))
            return (type_pref, year, month, max_eid_len, doi_len)
            
        # Sort so that highest priority is last
        comp.sort(key=get_priority)
        keep_key = comp[-1]
        del_keys = comp[:-1]
        
        # Merge all external IDs into keep_key
        all_eids = set(work_dict[keep_key].get("external_ids", []))
        for dk in del_keys:
            all_eids.update(work_dict[dk].get("external_ids", []))
        work_dict[keep_key]["external_ids"] = sorted(list(all_eids))
        
        # Delete duplicate entries
        for dk in del_keys:
            print(f"Merging {work_dict[dk]['title']} into {work_dict[keep_key]['title']}")
            del work_dict[dk]
            
    return work_dict


def find_preprint_repository(work_dict: Dict[str, Any]) -> Dict[str, Any]:
    """
    Fetches preprint metadata via requests with timeouts to populate the repository name.
    """
    for w in work_dict.values():
        if w.get("type") != "preprint":
            continue

        doi = w.get("doi", "")
        if doi:
            print(f"Trying to find host repository for article: {w['title']}")
            if "eLife" in doi:
                w["journal"] = "eLife"
            else:
                try:
                    # Timeout set to 5 seconds to prevent indefinite hangs
                    doi_data = requests.get(doi, timeout=5)
                    url = doi_data.url
                    parsed = urlparse(url)
                    netloc = parsed.netloc
                    if netloc.startswith("www."):
                        netloc = netloc[4:]
                    
                    parts = netloc.split(".")
                    if len(parts) >= 2:
                        domain = parts[-2]
                    else:
                        domain = netloc
                        
                    if "rxiv" in domain:
                        domain = domain.replace("rxiv", "Rxiv")
                    w["journal"] = domain
                except Exception as e:
                    print(f"Could not lookup preprint: {w['title']} ({e})")

    return work_dict


def load_funding(funding_path: str) -> Dict[str, Any]:
    """Loads a single funding record."""
    in_funding_dict = load_xml(funding_path)
    out_funding_dict = {
        "title": get_recursive_key(in_funding_dict, "funding:title", "common:title"),
        "role": get_recursive_key(in_funding_dict, "funding:organization-defined-type"),
        "org": get_recursive_key(in_funding_dict, "common:organization", "common:name"),
        "id": get_recursive_key(
            in_funding_dict,
            "common:external-ids",
            "common:external-id",
            "common:external-id-value",
        ),
        "start_year": get_recursive_key(
            in_funding_dict, "common:start-date", "common:year"
        ),
        "end_year": get_recursive_key(
            in_funding_dict, "common:end-date", "common:year"
        ),
        "value": get_recursive_key(in_funding_dict, "common:external-ids", "#text"),
    }
    return out_funding_dict


def load_review(review_path: str) -> Dict[str, Any]:
    """Loads a peer review record, lookup journal name by ISSN online."""
    in_review_dict = load_xml(review_path)
    issn = in_review_dict["peer-review:review-group-id"][5:]

    potential_name = ""
    try:
        r = requests.get(f"https://portal.issn.org/resource/ISSN/{str(issn)}", timeout=5)
        if r.status_code == 200:
            match = re.search(r"<title>ISSN\s+[\dXY-]+\s+-\s+(.*?)</title>", r.text, re.IGNORECASE)
            if match:
                potential_name = match.group(1).strip()
                if "(" in potential_name:
                    idx = potential_name.find("(")
                    potential_name = potential_name[:idx].strip()
    except Exception as e:
        logger.warning(f"Could not lookup ISSN {issn}: {e}")

    if not potential_name:
        print(f"Could not identify ISSN {issn}")

    out_review_dict = {
        "year": get_recursive_key(
            in_review_dict, "peer-review:review-completion-date", "common:year"
        ),
        "role": get_recursive_key(in_review_dict, "peer-review:review-type"),
        "org": potential_name.title(),
    }
    return out_review_dict


def folder_to_dict(path: str, load_fun: Callable[[str], Any]) -> Dict[str, Any]:
    """Reads all XML files in a directory and applies a loader function to each."""
    _dict = {}
    if not os.path.exists(path):
        logger.warning(f"Directory does not exist: {path}")
        return _dict
        
    xml_list = os.listdir(path)
    for x in xml_list:
        if x.endswith(".xml"):
            _dict[x[:-4]] = load_fun(os.path.join(path, x))
    return _dict


def extract_orcid_info(orcid_dir: str) -> Dict[str, Any]:
    """
    Coordinates XML parsing across personal, works, and affiliations,
    caching findings as an ORCID.json file.
    """
    json_path = os.path.join(orcid_dir, "ORCID.json")
    if os.path.isfile(json_path):
        print("Loading ORCID dict from local json.")
        with open(json_path, encoding="utf-8") as f:
            cached = json.load(f)

        # Caches written before the service section existed lack that key. Read
        # just that folder back in rather than re-parsing everything, which would
        # repeat the preprint and ISSN network lookups.
        if "service" not in cached:
            print("Adding service section to cached json.")
            cached["service"] = folder_to_dict(
                os.path.join(orcid_dir, "affiliations", "services"), load_affiliation
            )
            with open(json_path, "w", encoding="utf-8") as fp:
                json.dump(cached, fp, indent=4)

        return cached

    # Personal info
    person_path = os.path.join(orcid_dir, "person.xml")
    if not os.path.exists(person_path):
        raise FileNotFoundError(f"Missing required person.xml in {orcid_dir}")
        
    personal_info = load_xml(person_path)
    personal = {
        "lastname": get_recursive_key(
            personal_info, "person:name", "personal-details:family-name"
        ),
        "givenname": get_recursive_key(
            personal_info, "person:name", "personal-details:given-names"
        ),
        "links": {
            "ORCID": "https://orcid.org/"
            + get_recursive_key(personal_info, "person:name", "@path")
        },
    }
    
    personal["fullname"] = personal["givenname"] + " " + personal["lastname"]
    from orcid_cv.utils import initialize_name
    personal["name-short"] = initialize_name(personal["fullname"])
    
    first_space = personal["fullname"].find(" ")
    personal["firstname"] = personal["fullname"][0:first_space] if first_space != -1 else personal["fullname"]

    # Extract links
    if "researcher-url:researcher-urls" in personal_info.keys():
        link_list = get_recursive_key(
            personal_info,
            "researcher-url:researcher-urls",
            "researcher-url:researcher-url",
        )
        if isinstance(link_list, list):
            for link in link_list:
                personal["links"][link["researcher-url:url-name"]] = link["researcher-url:url"]
        elif isinstance(link_list, dict):
            personal["links"][link_list["researcher-url:url-name"]] = link_list["researcher-url:url"]

    # Get primary email
    emails_wrapper = personal_info.get("email:emails")
    if emails_wrapper and isinstance(emails_wrapper.get("email:email"), list):
        email_list = emails_wrapper["email:email"]
        primary_emails = [e["email:email"] for e in email_list if e.get("@primary") == "true"]
        if primary_emails:
            personal["email"] = primary_emails[0]
        elif email_list:
            personal["email"] = email_list[0]["email:email"]
        else:
            personal["email"] = ""
    elif emails_wrapper and isinstance(emails_wrapper.get("email:email"), dict):
        personal["email"] = emails_wrapper["email:email"].get("email:email", "")
    else:
        personal["email"] = ""

    # Parse XML folders to make dictionaries
    employment_dict = folder_to_dict(
        os.path.join(orcid_dir, "affiliations", "employments"), load_affiliation
    )
    education_dict = folder_to_dict(
        os.path.join(orcid_dir, "affiliations", "educations"), load_affiliation
    )
    service_dict = folder_to_dict(
        os.path.join(orcid_dir, "affiliations", "services"), load_affiliation
    )
    work_dict = folder_to_dict(os.path.join(orcid_dir, "works"), load_work)
    funding_dict = folder_to_dict(os.path.join(orcid_dir, "fundings"), load_funding)
    review_dict = folder_to_dict(os.path.join(orcid_dir, "peer_reviews"), load_review)

    # Check for duplicate work dicts & get preprint repositories
    work_dict = prune_duplicate_works(work_dict)
    work_dict = find_preprint_repository(work_dict)

    out_dict = {
        "personal": personal,
        "work": work_dict,
        "employment": employment_dict,
        "education": education_dict,
        "service": service_dict,
        "funding": funding_dict,
        "reviews": review_dict,
    }

    # Save cache
    print("Saving local json.")
    with open(json_path, "w", encoding="utf-8") as fp:
        json.dump(out_dict, fp, indent=4)

    return out_dict
