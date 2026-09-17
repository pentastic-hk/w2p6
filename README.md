# Word Docx Tables Screenshot Taker

This python script automatically takes screenshots for every table in the word docx.
Multiple rows will be merged together if their visible height is less than a certain number or rows.

The free software *LibreOffice* is used to render the Word Docx.
You're highly recommended to use this rendering engine
to obtain a nearly identical screenshot as if you take it in MS Word.

See [INSTALL.md](./INSTALL.md) for instructions to install this project.

## How to Use

```bash
python landsdetail2img.py "QC FUP v1.2.docx"
#    -> writes into a sibling folder "QC FUP v1.2-screenshots/"
```

## Smart Features

### Auto detection of image height

This script has a line count detection that decides
whether multiple consecutive rows should be merged into 1 image.

Currently, that default number is 25 visible lines.
That means, 2 consecutive rows won't be merged into 1 image
if their total line count exceed 25 lines.

This could be overriden by setting the flag `--max-rows`.

## The python script filename

It is just a historical reason.

Originally it is planned to receive the generated landscape-detail format
from the x2w1 project, but it turns out that this script is more powerful than that.
