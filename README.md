
# pptx-to-md

Two-part converter from PowerPoint/Impress to beamer-friendly,
Pandoc-style markdown. 
(1) convert pptx (or ppt, or odp) into intermediate
YAML, then (2) convert YAML into markdown.

## pptx-to-yaml usage

```bash
$ ./pptx-to-yaml.py [--use-server HOST:PORT] INPUT_FILE OUTPUT_FILE IMAGE_DIR
```

`INPUT_FILE` is the path to some ppt, pptx or Impress file.

`OUTPUT_FILE` is the name of the YAML file to be written.

IMAGE_DIR is a directory where images will be extracted to.

It will be created if it doesn't exist.

## PowerPoint features supported

Title text, "outline" text (i.e. bullets) and embedded
graphics like JPEGs or PNGs are handled reasonably well.

Embedded "metafiles" (EMF or WMF vector graphics)
should get converted to SVG. (And thence to EPS, if you
use `convert.sh`.)

If it finds any tables, drawing shapes (arrows/boxes etc),
`pptx-to-yaml.py` tries to collect them all together
and export them as an SVG.

## soffice server

`pptx-to-yaml.py` attempts to start an `soffice` process
and communicate with it over port 2002 on the local host;
it's the `soffice` process that "knows" how to read
PowerPoint etc files.

However, the `HOST:PORT` can be supplied if you prefer
to run your own instance of `soffice` as a separate process.
Which you might want to, since:

a.  If you have a lot of files to convert, you can just
    keep one `soffice` process running, and re-use it,
    avoiding the time taken to start a new process for
    each document.
b.  Sometimes `pptx-to-yaml.py` just doesn't seem to
    start the `soffice` process up correctly - I have no idea why.

So you could start the server process using something
like the following:

```bash
$ xterm -e 'soffice --accept="socket,host=localhost,port=2002;urp;" \
    --norestore --nologo --nodefault --headless' &
```

... which will open an `soffice` instance running in its own terminal
window; and then specify `HOST` and `PORT` to `pptx-to-yaml.py`.


## yaml-to-md usage

```bash
$ ./yaml-to-md.py INPUT_FILE OUTPUT_FILE
```

Just takes an input file and output file.

## utility script - convert.sh

usage:

```bash
$ ./convert.sh [INPUT_FILE..]
```

Convenience wrapper around pptx-to-yaml and yaml-to-md. Also converts
SVG files to encapsulaed PostScript (EPS) for use by LaTeX,
and attempts to use Pandoc to output .tex and .pdf.
(If it fails, the .md file needs some tidying, so the
.pdf just isn't output.)


## prerequisites

For `pptx-to-yaml.py` and `yaml-to-md.py`:

-   Python 3.5 or greater
-   LibreOffice 5.1.6. On Ubuntu 16.04 (xenial), this
    can be installed with `sudo apt-get install libreoffice`.
-   `python3-uno`. On Ubuntu, this can be installed with
    `sudo apt-get install python3-uno`.
-   pyyaml. Most easily installed with something like
    `pip3 install --user pyyaml`.

For `convert.sh`:

-  Requires bash, sed, [Inkscape](https://inkscape.org)
  (for converting SVG to EPS) and [Pandoc](https://pandoc.org/)
  (for converting .md to .tex or .pdf).

## idiosyncracies

Exported/graphics files are all referred to by absolute pathname,
so if you want to move your generated files around,
you'll have to edit any references to them in the
YAML/markdown, as appropriate.

## Portability

Not at all portable, and not tested on any other platform
other than Ubuntu 16.04, nor with any other version of
LibreOffice than 5.1.6.

## Reporting bugs

You can if you want, but there's no guarantee I'll fix them.
The scripts are really just offered as a starting point for
anyone else who wants to improve them.

## troubleshooting

### can't connect

If you get some error saying `pptx-to-yaml.py` couldn't connect to the
server -- kill any stray soffice process and try again.

If it still fails, possibly add a bigger `time.sleep`
in the script, or just run your own server process.

### can't open wmf/emf files in inkscape

Try opening them in lodraw.

### Pandoc/LaTeX fails to compile .pdf

Lots of things could have gone wrong. By default, Pandoc
uses pdflatex, which will choke on many Unicode symbols.
Graphics might not have converted. etc.

The only thing to do is take a look at the original
PowerPoint file, and the generated markdown, and see if you
can fix whatever went wrong.



