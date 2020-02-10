#!/usr/bin/env python3

"""
intermediate yaml to markdown conversion
"""

import sys
import yaml

def yaml_to_markdown(yaml, outfile):
  """Given a list of dicts representing PowerPoint slides
  -- presumably loaded from a YAML file -- convert to
  markdown and print the result on the file-like
  object 'outfile'.
  """

  for slide in yaml:
    slide_to_markdown(slide, outfile)

def get_title(slide):
  """return title or None. Deletes title from dict"""
  shapes = slide["conts"]
  found = False
  for i, shape in enumerate(shapes):
    if shape["ShapeType"] == "com.sun.star.presentation.TitleTextShape":
      found = True
      title = shape
      break
  if found:
    del shapes[i]
    return title["String"].replace("\n", " ")

def slide_to_markdown(slide, outfile):
  shapes = slide["conts"]
  title = get_title(slide)
  if not title:
    title = "SLIDE"
  print("### " + title + "\n", file=outfile)

  for shape in shapes:
    if shape["ShapeType"] == "com.sun.star.drawing.GraphicObjectShape":
      add_graphic(shape, outfile)
    # all Groups should've been converted to SVG
    elif shape["ShapeType"] == "com.sun.star.drawing.GroupShape":
      print("grouping ...\nslide title: ", title)
      add_graphic(shape, outfile)
    elif shape["ShapeType"] == "com.sun.star.presentation.TitleTextShape":
      out_str = "(TABLE not converted from PowerPoint)"
      print(out_str + "\n", file=outfile)
    elif "elements" in shape:
      add_list(shape, outfile)
    elif "String" in shape and shape["String"]:
      add_text(shape, outfile)
    else:
      out_str = "<!-- sl: %(slideNum)s, shp: %(shapeNum)s, type: %(shapeType)s !-->" % {
                    "slideNum" : slide["slideNum"],
                    "shapeNum" : shape["shapeNum"],
                    "shapeType" : shape["ShapeType"] }
      print(out_str + "\n", file=outfile)

def add_text(shape, outfile):
    """
    convert a text-like Shape to a string, and
    print to 'outfile'
    """

    print( shape["String"].strip() + "\n", file=outfile)

def add_list(shape, outfile):
  """
  Given a shape that represents an 'Outline' --
  OpenOffice's representation of a bulleted or numbered
  list -- attempt to convert the elements into
  a sensible Markdown list, and write to
  "outfile".
  """

  els = shape["elements"]

  indent = 0

  def item_to_str(item):
    s = (' ' * indent * 4) + "-   " + item["String"].strip()
    return s

  # handle first item

  output = [item_to_str(els[0])]

  def dump_output():
    print( "\n".join(output) + "\n", file=outfile)

  if len(els) == 1:
    dump_output()
    return

  # handle rest of items

  last_el = els[0]

  for el in els[1:]:
    # int-ify the level if None
    if el["NumberingLevel"] is None:
      el["NumberingLevel"] = 0
    if last_el["NumberingLevel"] is None:
      last_el["NumberingLevel"] = 0

    # new indent
    if el["NumberingLevel"] > last_el["NumberingLevel"]:
      indent += 1
    elif el["NumberingLevel"] < last_el["NumberingLevel"]:
      indent = max(0, indent-1)
    else:
      pass

    #print("  new indent:", indent)

    if len(el["String"]) > 1:
      output.append(item_to_str(el))

    last_el = el

  dump_output()

def add_graphic(shape, outfile):
  """
  Given a Shape representing some graphics object
  (e.g. jpg, png, MetaFile, SVG), write out
  the markdown to show it on "outfile".
  """

  if "String" in shape and shape["String"]:
    alt_text = shape["String"]
  else:
    alt_text = ""

  if "exported_svg_filename" in shape:
    filename = shape["exported_svg_filename"]
  else:
    filename = shape["exported_filename"]

  link = "![%(alt_text)s](%(filename)s)"  % { "alt_text" : alt_text,
                                              "filename" : filename }

  print(link + "\n", file=outfile)

# typical image types:
#  image/jpeg, image/png, image/gif

# text shapes:
#   TextShape, NotesShape, SubtitleShape, OutlinerShape,
#   TitleTextShape, ?CustomShape, possibly ?RectangleShape
def convert_file(input_file, output_file):
  """start an soffice server, then convert input file to output file
    using image dir."""

  with open(input_file, "r") as input:
    y = yaml.load(input, Loader=yaml.SafeLoader)
    with open(output_file, "w") as output:
      yaml_to_markdown(y, output)

MAIN="__main__"
#MAIN=None

def main():
  """main"""
  args = sys.argv[1:]

  if len(args) != 2:
    print("usage: pptx-to-md.py INPUT_FILE OUTPUT_FILE")
    sys.exit(1)

  input_file, output_file = args
  convert_file(input_file, output_file)


if __name__ == MAIN:
  main()

