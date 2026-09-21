"""项目初始化环境的冒烟测试。"""

import sys

import pptx
import pydantic
import pydantic_settings

import ppt_maker


def test_python_and_required_packages_are_available() -> None:
    """验证项目固定在 Python 3.12 且第一阶段依赖可以导入。"""
    assert sys.version_info[:2] == (3, 12)
    assert ppt_maker.__version__ == "0.1.0"
    assert pydantic.__version__
    assert pydantic_settings.__version__
    assert pptx.__version__
