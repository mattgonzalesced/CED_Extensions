# coding: utf8
from Autodesk.Revit import Exceptions
from pyrevit import revit, DB, forms


class Workset:
    def __init__(self, name):
        self.name = name

    @staticmethod
    def read_from_txt():
        file = forms.pick_file(file_ext="txt", title="Select Text File")
        workset_list = []
        with open(file, "r") as text:
            for line in text.readlines():
                workset_list.append(Workset(line.strip("\n")))

        return workset_list

    def create(self, doc=revit.doc):
        try:
            DB.Workset.Create(doc, self.name)
        except Exceptions.ArgumentException as error:
            print(
                "Failed to create a workset named {}\n{}".format(
                    self.name, error.Message
                )
            )

    def save_to_file(self):
        # TODO
        return
