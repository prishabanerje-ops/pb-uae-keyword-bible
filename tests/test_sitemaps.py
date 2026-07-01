import pytest
from keyword_bible import sitemaps

XML = """<?xml version='1.0'?>
<urlset xmlns='http://www.sitemaps.org/schemas/sitemap/0.9'>
<url><loc>https://www.policybazaar.ae/health-insurance/</loc></url>
<url><loc>https://www.policybazaar.ae/daman-health-insurance-hipvd/</loc></url>
</urlset>"""

def test_parse_extracts_locs():
    urls = sitemaps.parse(XML)
    assert urls == [
        "https://www.policybazaar.ae/health-insurance/",
        "https://www.policybazaar.ae/daman-health-insurance-hipvd/",
    ]

def test_parse_empty():
    assert sitemaps.parse("<urlset></urlset>") == []
