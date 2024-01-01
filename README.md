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