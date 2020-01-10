#!/usr/bin/env python3
import sys
from fbs_runtime.application_context.PyQt5 import ApplicationContext, cached_property

from window import MainWindow


class AppContext(ApplicationContext):
    def run(self):
        self.window.show()
        return self.app.exec_()

    def get_ui(self):
        return self.get_resource('main.ui')

    @cached_property
    def window(self):
        return MainWindow(self)


if __name__ == '__main__':
    context = AppContext()
    exit_code = context.run()
    sys.exit(exit_code)
