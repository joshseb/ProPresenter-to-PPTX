"""PyInstaller hook for google-protobuf (namespace package fix)."""
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Collect every google.protobuf submodule (descriptor, descriptor_pool, etc.)
hiddenimports = collect_submodules('google.protobuf')

# Include any .py data files shipped with the package
datas = collect_data_files('google.protobuf')
