# -*- coding: utf-8 -*-
class MEPKitError(Exception): pass
class NotFound(MEPKitError): pass
class InvalidInput(MEPKitError): pass
class RevitOpFailed(MEPKitError): pass