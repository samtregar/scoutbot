#!/bin/env python
def load_src(name, fpath):
    import os, imp
    return imp.load_source(name, os.path.join(os.path.dirname(__file__), fpath))
load_src("ScoutBot", "../ScoutBot.py")
from ScoutBot import ScoutBot
ScoutBot().slackbot()
