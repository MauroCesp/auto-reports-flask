# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

from flask import Blueprint

# The first argument, "home_blueprint", is the Blueprint’s name, which is used by Flask’s routing mechanism. 
# The second argument, __name__, is the Blueprint’s import name, which Flask uses to locate the Blueprint’s resources.
# The third argument is url_prefix: the path to prepend to all of the Blueprint’s URLs
blueprint = Blueprint(
    'home_blueprint',
    __name__,
    url_prefix=''
)
