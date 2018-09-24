import collections
from pathlib import Path
from typing import Union
import tempfile
import os


def is_iterable(obj) -> bool:
    return not isinstance(obj, (str, bytes)) and isinstance(obj, collections.Iterable)


def safe_path(path: Union[str, Path]) -> Path:
    path = path if isinstance(path, Path) else Path(path)
    counter = 0
    new_path = path
    while new_path.exists():
        counter += 1
        new_path = new_path.with_name(
            '{stem} ({counter:d})'.format(stem=new_path.stem, counter=counter)
        ).with_suffix(*new_path.suffixes)
    return new_path


def temp_path(suffix) -> Path:
    handle, path = tempfile.mkstemp(suffix)
    os.close(handle)
    return path