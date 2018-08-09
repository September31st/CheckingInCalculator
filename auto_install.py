# -*- coding: utf-8 -*-
import os


def import_tools():
    try:
        import xlrd
    except:
        print "install xlrd.."
        os.system("pip install xlrd")
        import xlrd
        from xlrd import xldate_as_tuple
        print "xlrd installed.."
    try:
        import xlwt
    except:
        print "install xlwt.."
        os.system("pip install xlwt")
        import xlwt
        print "xlwt installed.."
    try:
        import datetime
    except:
        print "install datetime.."
        os.system("pip install datetime")
        import datetime
        from datetime import datetime, timedelta
        print "datetime installed.."
