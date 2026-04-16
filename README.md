# Thesis DOCX Format Checker

This project reads the thesis-format PDF specification and checks uploaded `.docx` files against the parts that can be inspected programmatically.

It reports:

- what is out of spec
- where the problem appears
- why it is a problem
- how to fix it

## Current coverage

- A4 page size
- margins: top 2.5 cm, bottom 2.5 cm, left 3 cm, right 2 cm
- chapter heading format
- section heading format
- abstract and keyword checks
- body font and body size checks
- 1.5 line spacing checks
- page-number field detection
- watermark-object detection
- Word protection detection

## Known limits

- DOCX analysis cannot fully reconstruct final pagination.
- Watermark checks only detect watermark-like objects in the package.
- Theme- or style-inherited fonts may not always appear at run level.
- Paper weight, duplex printing, and final PDF security still need manual review.

## Run locally with Streamlit

```bash
streamlit run app.py
```

## Streamlit Cloud

- Set the main file to `app.py`
- Deploy normally on Streamlit Cloud
- Do not run this project with `python app.py` on Streamlit Cloud

## Alternative local Flask version

The old Flask UI has been replaced by a Streamlit entrypoint so deployment on Streamlit Cloud works correctly.

## Files

- `app.py`: Streamlit upload UI
- `thesis_checker.py`: DOCX analysis engine
- `requirements.txt`: dependencies
