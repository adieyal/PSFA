import os
import sys
from pbs import inkscape

code_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)))
project_root = os.path.join(code_dir, os.path.pardir)
output_dir = os.path.join(project_root, "output")
for path, _, files in os.walk(output_dir):
    for f in files:
        if f.endswith(".svg"):
            in_path = os.path.join(path, f)
            out_path = os.path.join(path, f.replace(".svg", ".pdf"))
            print "Converting: %s -> %s" % (in_path, out_path)
            inkscape("-A", out_path, in_path)

            
