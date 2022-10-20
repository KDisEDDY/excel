# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import os
import sys

import FollowTemplateFunction


def app_path():
    """Returns the base application path."""
    if hasattr(sys, 'frozen'):
        # Handles PyInstaller
        return os.path.dirname(sys.executable)  # 使用pyinstaller打包后的exe目录
    return os.path.dirname(__file__)  # 没打包前的py目录


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    followTemplate = FollowTemplateFunction.EUFollowTemplateFunc()
    followTemplate.traverseCurrentDirectory(app_path())
