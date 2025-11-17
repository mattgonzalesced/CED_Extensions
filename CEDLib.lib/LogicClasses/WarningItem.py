# -*- coding: utf-8 -*-


class WarningItem(object):
    def __init__(self, code=None, message=None):
        self._code = code
        self._message = message

    def get_code(self):
        return self._code

    def set_code(self, value):
        self._code = value

    def get_message(self):
        return self._message

    def set_message(self, value):
        self._message = value

    # Helpers
    def to_dict(self):
        return {
            "code": self._code,
            "message": self._message,
        }
