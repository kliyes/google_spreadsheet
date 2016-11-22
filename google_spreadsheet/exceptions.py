# encoding=utf8
'''
Created on 2016-11-10

@author: jingyang <jingyang@nexa-corp.com>
'''


class APIException(Exception):
    """
    Base class of all exceptions
    """
    status_code = 400

    def __init__(self, error):
        self.messages = error

    def __repr__(self):
        return "{self.status_code}: {self.messages}".format(self=self)

    __str__ = __repr__


class NotFound(APIException):
    """
    Trying to open a non-existent spreadsheet.
    """
    status_code = 404


class PermissionDenied(APIException):
    """
    Trying to open a inaccessible spreadsheet.
    """
    status_code = 403


class BadRequest(APIException):
    """
    Bad request exceptions.
    """


class IncorrectCellLabel(APIException):
    """
    Trying to convert an incorrect cell label
    """
