# ORCID_cv
Pipeline for creating a CV from the ORCID xml dump file.\
To download your own xml dump file login to your ORCID account then:\
User -> Account Settings -> "Download your ORCID data"\
Follow the 'Quick build' example in ```examples.py```:
```python
import orcid_cv as ocv
orcid_dir = r"C:\Users\somlab\Downloads\0000-0002-6806-3302"
output_fname = r"C:\Users\somlab\Downloads\quick_build_output_example.pdf"
ocv.quick_build(orcid_dir, output_fname)
```
This will generate something like [this example pdf](quick_build_output_example.pdf).

## Choosing a PDF engine
The same parsed ORCID data can be typeset by either of two backends:

| backend | engine | notes |
| --- | --- | --- |
| `reportlab` (default) | [reportlab](https://pypi.org/project/reportlab/) | flowables and tables, drawn directly |
| `typst` | [typst](https://pypi.org/project/typst/) | generates Typst markup and compiles it; the pip package bundles the compiler, so nothing extra to install |

Pass `backend` to `quick_build`:
```python
ocv.quick_build(orcid_dir, output_fname, backend='typst')
```
or to `make_document_config` when assembling a CV section by section. Every
`add_*_section` call and `build_document` dispatch on the config, so only that one
line changes:
```python
config = ocv.make_document_config('greenspon-default', backend='typst')
config['page_footer'] = True  # 'Page X of Y' + today's date in the footer
elements = []
ocv.add_person_section(elements, orcid_dict, config)
ocv.add_affiliation_section(elements, orcid_dict, config, 'Employment', 'employment')
ocv.add_work_section(elements, orcid_dict, config, 'Research Publications', 'journal-article')
ocv.build_document(output_fname, elements, config, title='My CV')
```
The two engines produce very similar, not identical, output: typst typesets a little
more compactly and handles unusual characters in titles more gracefully, while
reportlab strips anything that looks like an HTML tag.

To keep the generated Typst markup (to tweak it by hand or compile it yourself),
pass `save_source`:
```python
ocv.build_document(output_fname, elements, config, save_source=r"cv.typ")
```

## Layout of the package
* `parser.py` – reads the ORCID XML dump into a dictionary (cached as `ORCID.json`)
* `content.py` – turns that dictionary into markup-free entries shared by both backends
* `builder.py` / `styles.py` – reportlab document assembly and styling
* `typst_builder.py` / `typst_styles.py` – the same, emitting Typst markup
* `config.py` – per-style, per-backend settings (fonts, sizes, spacing)

## Environment
```bash
conda env create -f environment.yml
conda activate orcid_cv
```
