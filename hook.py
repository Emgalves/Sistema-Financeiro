def pre_init_hook():
    import os
    import sys
    
    if getattr(sys, 'frozen', False):
        application_path = sys._MEIPASS
        os.environ['PYTHONPATH'] = os.path.join(application_path, 'src')
        sys.path.insert(0, os.path.join(application_path, 'src'))
        sys.path.insert(0, os.path.join(application_path, 'src', 'config'))
