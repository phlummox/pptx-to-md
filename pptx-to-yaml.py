#!/usr/bin/env python3

"""
convert Impress slides to yaml, thence to be turned
into markdown or similar.
"""

import subprocess
from multiprocessing import Process
import os
from os.path import abspath
import sys
import time
import uno
import yaml

def launch_office(host, port):
  """launch soffice, but not in _background_. (so we omit --headless)"""

  args = ['soffice', '--accept=socket,host=%s,port=%s;urp;' % (host,port),
          '--norestore',  '--nologo', '--nodefault', '--headless']

  subprocess.run(args, check=True)

class OfficeServer:
  """
  wrapper to do bracketing of soffice process
  """

  def __init__(self, host, port):
    self.host = host
    self.port = port
    self.process = None

  def __enter__ (self):
      # Code to start a new transaction
      print("launching office, args = ", (self.host, self.port))

      self.process = Process(target=launch_office, args=(self.host, self.port))
      self.process.start()
      return self.process

  def close(self):
    self.process.terminate()

  def __exit__ (self, _type, _value, _tb):
    # if tb is not None, some sort of exception occurred.
    # but we don't care
    self.close()




def get_connection_url(hostname, port, pipe=None):
  "openoffice url, given hostnane and port"

  if pipe:
      conn = 'pipe,name=%s' % pipe
  else:
      conn = 'socket,host=%s,port=%d' % (hostname, port)
  return 'uno:%s;urp;StarOffice.ComponentContext' % conn




class DesktopWrapper:
  """wrapper to do bracketing of Uno Desktop object. i.e. withDesktop"""

  def __init__(self, host, port):
    self.host = host
    self.port = port
    url = get_connection_url(host, port)
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    resolver = smgr.createInstanceWithContext(
        'com.sun.star.bridge.UnoUrlResolver', ctx)
    self.remote_context = remote_context = resolver.resolve(url)
    self.serviceManager = remote_smgr = remote_context.ServiceManager
    self.desktop = remote_smgr.createInstanceWithContext(
        'com.sun.star.frame.Desktop', remote_context)


  def __enter__ (self):
      return self

  def __exit__ (self, _type, _value, _tb):
    # if tb is not None, some sort of exception occurred.
    # but we don't care

    #self.desktop.terminate()
    pass

  def createUnoService( self, cClass ):
      """A handy way to create a global objects within the running OOo.
      """
      return self.serviceManager.createInstance( cClass )

  def openDocument( self, path ):
    """open a doc given path. turns it into a URL acceptable to libreoffice."""

    url = uno.systemPathToFileUrl(path)
    return self.desktop.loadComponentFromURL( url,"_blank", 0, () )

  def createInstanceWithContext(self, className ):
    """
    utility func to do serviceManager.createInstanceWithContext.

    example:

    >>> dwrap = DesktopWrapper("localhost", 2020)
    >>> dwrap.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter")
    """
    return self.serviceManager.createInstanceWithContext( className, self.remote_context)

def plausible_file_extension(s):
  """given some string, likely a mime-type,
  return either a .SOMETHING extension, or an
  empty string"""

  ext = ""

  l = len(s)
  for i in range(l-1, -1, -1):
    if s[i].isalpha():
      ext = s[i] + ext
    else:
      break
  return ext

def mk_image_filename(slideNum,shapeNum, mimeType):
  """
  Given slideNum (number of slide within doc; ideally,
  starting from 1),
  shapeNum (number of shape within slide; ideally, starting
  from 1)
  and mimeType, return appropriate (base) filename. (No dir)
  """

  graphic_basename = "graphic-s%03d-g%03d" % (slideNum,shapeNum)
  graphic_ext = plausible_file_extension(mimeType)
  if len(graphic_ext) > 0:
    graphic_filename = graphic_basename + "." + graphic_ext
  else:
    graphic_filename = graphic_basename
  return graphic_filename

def export_to_graphics_file(dwrap, shape, abs_filename, mimeType):
  """
  Export a graphics shape from a file, to an external file.

  dwrap: dwrap object
  shape: shape to export
  abs_filename: absolute path f where to export it
  mimeType: mime type of graphic
  """

  exporter = dwrap.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter")
  exporter.setSourceDocument(shape)

  # make an arg list for exporter
  arg1 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
  arg1.Name, arg1.Value = "URL", uno.systemPathToFileUrl(abs_filename)

  arg2 = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
  arg2.Name, arg2.Value = "MediaType", mimeType

  exporter.filter([arg1,arg2])

#SHAPES = []

class SlideConverter:
  """given some file we want to convert to yaml, do so".

  An store a few useful bits of data so we don't
  have to pass them around."""

  def __init__(self, dwrap, filepath):
    """dwrap = a dwrap object,
    filepath = path to document to convert,
    """

    self.dwrap = dwrap
    self.filepath = filepath

    self.doc = doc = dwrap.openDocument( abspath(filepath) )
    self.controller = doc.getCurrentController()
    self.dispatcher = dwrap.createUnoService("com.sun.star.frame.DispatchHelper")

  def glomShapesTogether(self, page):
    """
    glom a bunch of shapes together that are probably part of
    a drawing or table or diagram, and convert them into a
    metafile.

    The newly-created metafile can then, hopefully, be
    converted/exported like any other graphic.
    """

    def isDrawLike(shape):
      return shape.ShapeType in ["com.sun.star.drawing.OLE2Shape",
                                 "com.sun.star.drawing.LineShape",
                                 "com.sun.star.drawing.CustomShape",
                                 "com.sun.star.drawing.TableShape",
                                 "com.sun.star.drawing.GroupShape"]

    # shapes that are probably part of a drawing or
    # diagram of some sort
    drawingIshShapes = []

    for j in range(0, page.Count):
      shape = page.getByIndex(j)
      if isDrawLike(shape):
        drawingIshShapes.append(shape)

    if len(drawingIshShapes) == 0:
      return

    print("Grouping", len(drawingIshShapes), "probably-drawing shapes")

    # build up a collection of all shapes on page
    shapeColl = self.dwrap.createUnoService("com.sun.star.drawing.ShapeCollection")
    for drawingShape in drawingIshShapes:
      shapeColl.add(drawingShape)
    # make them a Group
    shapeGroup = page.group(shapeColl)
    # convert group into a metafile
    self.controller.select(shapeGroup)
    self.dispatcher.executeDispatch(self.controller.Frame, ".uno:ConvertIntoMetaFile", "", 0, [])

  def convert(self, image_dir):
    """
    image_dir = where to write image files to.

    returns a list of dicts (one for each slide), suitable for YAMLizing.

    TODO: ought to grab document metadata as well.
    """

    filepath = self.filepath
    doc = self.doc
    dwrap = self.dwrap

    # will be returned.
    output_slides = []

    # map from url to saved files
    exported_images = {}

    print("processing file", filepath)
    for i, page in enumerate(doc.DrawPages):

      print("[-] page", i)

      slideObj = { "slideNum" : i }
      output_slides.append(slideObj)

      slideConts  = list([])

      # first of all, glom together any OLE shapes/
      # things that look like part of a table/chart/drawing
      # into a single MetaFile for easy exporting.
      self.glomShapesTogether(page)

      for j in range(0, page.Count):

        shape = page.getByIndex(j)

        try:
          shape_String = shape.String
        except Exception as _ex:
          shape_String = None

        stuff = { "ShapeType": shape.ShapeType,
                  "String": shape_String,
                  "shapeNum": j
                  }


        if shape.ShapeType == "com.sun.star.drawing.GraphicObjectShape":
          # do we need all these? Not really.
          stuff["graphic"] = True
          stuff["GraphicURL"] = shape.GraphicURL
          stuff["GraphicStreamURL"] = shape.GraphicStreamURL
          stuff["Graphic.MimeType"] = shape.Graphic.MimeType
          stuff["Graphic.Size.Width"] = shape.Graphic.Size.Width
          stuff["Graphic.Size.Height"] = shape.Graphic.Size.Height

          imgUrl = shape.GraphicURL

          if imgUrl in exported_images:
            graphic_filename = exported_images[imgUrl]
            abs_filename = os.path.abspath(image_dir) + "/" + graphic_filename
            stuff["exported_filename"] = abs_filename

          else:
            # we need to export image and add to dict

            if not os.path.exists(image_dir):
              os.mkdir(image_dir)

            mimeType = shape.Graphic.MimeType
            graphic_filename = mk_image_filename(i+1,j+1, mimeType)
            abs_filename = os.path.abspath(image_dir) + "/" + graphic_filename
            print("exporting to", abs_filename)
            export_to_graphics_file(dwrap, shape, abs_filename, mimeType)
            stuff["exported_filename"] = abs_filename
            exported_images[imgUrl] = graphic_filename

            if "wmf" in mimeType or "emf" in mimeType:
              print("making SVG as well")
              graphic_filename = mk_image_filename(i+1,j+1, "svg")
              abs_filename = os.path.abspath(image_dir) + "/" + graphic_filename
              export_to_graphics_file(dwrap, shape, abs_filename, "image/svg+xml")
              stuff["exported_svg_filename"] = abs_filename
        elif shape.ShapeType == "com.sun.star.drawing.OLE2Shape":
          raise Exception("all OLE shapes should've been globbed!")
        elif shape.ShapeType == 'com.sun.star.drawing.TableShape':
          raise Exception("all TableShape should've been globbed!")
        # Any GroupShapes should've been globbed together
        # into a single metafile shape.
        elif shape.ShapeType == "com.sun.star.drawing.GroupShape":
          graphic_filename = mk_image_filename(i+1,j+1, "svg")
          abs_filename = os.path.abspath(image_dir) + "/" + graphic_filename
          print("Got groupshape, exporting to", abs_filename)
          export_to_graphics_file(dwrap, shape, abs_filename, "image/svg+xml")
          stuff["exported_svg_filename"] = abs_filename

        elif shape.hasElements():
          stuff["hasElements"] = True
          stuff["elementType"] = elType = shape.getElementType().typeName

          if elType == 'com.sun.star.text.XTextRange':
            it = shape.createEnumeration()
            els = []
            while it.hasMoreElements():
              els.append(it.nextElement())

            if len(els) > 1:
              stuff["numEls"] = len(els)
              stuff["elements"] = list()

              for el in els:
                elDict = { 'String' : el.String,
                           'NumberingLevel' : el.NumberingLevel,
                           'NumberingIsNumber': el.NumberingIsNumber,
                           'NumberingStartValue': el.NumberingStartValue }
                stuff["elements"].append(elDict)


        slideConts.append(stuff)

      slideObj["conts"] = slideConts

    doc.close(0) # closes without asking for save

    return output_slides

# typical image types:
#  image/jpeg, image/png, image/gif

# typical shapes we might encounter:
#   TextShape, NotesShape, SubtitleShape, OutlinerShape,
#   TitleTextShape, CustomShape, possibly RectangleShape

def convert_file(input_file, output_file, image_dir, use_server=None):
  """start an soffice server if requested, then convert input
    file to output file using image dir."""

  if use_server is None:
    host, port = ("localhost", 2002)
    s = OfficeServer(host, port)
    print("waiting for server..")
    for _i in range(0,6):
      print(".", end="")
      sys.stdout.flush()
      time.sleep(1)
  else:
    host, port = use_server.split(":")

  with DesktopWrapper(host, int(port)) as dwrap:
    converter = SlideConverter(dwrap, input_file)
    converted_slides = converter.convert(image_dir)
    with open(output_file, "w") as output:
      yaml.dump(converted_slides, output)

  if use_server is None:
    s.close()

MAIN="__main__"
#MAIN=None

def usage():
    """print usage message"""
    print("usage: pptx-to-md.py [--use-server HOST:PORT] INPUT_FILE OUTPUT_FILE IMAGE_DIR")

def main():
  """main"""
  args = sys.argv[1:]

  if len(args) not in [3,5]:
    usage()
    sys.exit(1)

  if len(args) == 5:
    if args[0] != "--use-server":
      usage()
      sys.exit(1)
    use_server = args[1]
    args.pop(0)
    args.pop(0)
  else:
    use_server = None

  input_file, output_file, image_dir = args
  convert_file(input_file, output_file, image_dir, use_server)

if __name__ == MAIN:
  main()

