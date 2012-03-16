import os
import sys
#from pbs import inkscape
import pbs

project_root = os.path.realpath(".")

inkscape = pbs.Command("c:\\Program Files\\Inkscape\\inkscape.exe")
for path, _, files in os.walk(os.path.join(project_root, "output")):
    for f in files:
        if f.endswith(".svg"):
            in_path = os.path.join(path, f)
            out_path = os.path.join(path, f.replace(".svg", ".pdf"))
            print "Converting: %s -> %s" % (in_path, out_path)
            inkscape("-A", '"%s"' % out_path, '"%s"' % in_path)

            
