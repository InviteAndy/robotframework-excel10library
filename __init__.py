#!/usr/bin/env python


#  Copyright 2018 MyInvite.nl.
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.
from Excel10Library import Excel10Library
from version import VERSION

_version_ = VERSION


class Excel10Library(Excel10Library):
    """
    This library provides keywords to allow
    basic control over Excel10 (xlsx) files
    from Robot Framework.
    """
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'
