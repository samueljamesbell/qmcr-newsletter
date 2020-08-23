# QMCR Newsletter Generator

Use this script to generate the QMCR newsletter.

## Overview

The generator will:

1. Fetch Google calendar events for the events bulletin
2. Fetch Google calendar events for the sports bulletin
3. Parse a CSV of approved bulletin entries
4. Generate a newsletter as a `.docx` file, using the provided template
5. Open the newsletter in Word for editing

You'll need to tweak the formatting (particularly the line breaks) in Word
after the newsletter is generated, before copying and pasting into your
favourite email client.

## Requirements

* Python 3.7.x
* A google `credentials.json` file, provided separately
* The dependendencies listed in `requirements.txt`

```
pip install -r requirements.txt
```

## Usage

```
python generate-newsletter.py path-to-bulletin-csv
```
