import os
import sys
from pbs import inkscape

for path, _, files in os.walk("output"):
    for f in files:
        if f.endswith(".svg"):
            in_path = os.path.join(path, f)
            out_path = os.path.join(path, f.replace(".svg", ".pdf"))
            print "Converting: %s -> %s" % (in_path, out_path)
            inkscape("-A", out_path, in_path)

            
