#!/usr/bin/env python3

import sys

from fbs_runtime.application_context.PyQt5 import ApplicationContext
from PyQt5 import uic


class AppContext(ApplicationContext):

    def run(self):
        self.setup_ui()
        self.window.show()
        return self.app.exec_()

    def setup_ui(self):
        ui = self.get_resource('main.ui')
        self.window = uic.loadUi(ui)


if __name__ == '__main__':
    app = AppContext()
    exit_code = app.run()
    sys.exit(exit_code)
