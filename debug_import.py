from importlib import import_module
m = import_module('src.gui_fullscreen_match')
print('module', m)
print('has FullScreenMatchGUI:', hasattr(m, 'FullScreenMatchGUI'))
cls = getattr(m, 'FullScreenMatchGUI', None)
print('FullScreenMatchGUI:', cls)
import inspect
print('isclass:', inspect.isclass(cls))
print('bases:', getattr(cls, '__mro__', None))
