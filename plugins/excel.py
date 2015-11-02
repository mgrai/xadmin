# coding:utf-8
#author 'sai'
#eg : import_excel = True

import xadmin
from django.conf import settings
from xadmin.views import BaseAdminPlugin, ListAdminView
from django.template import loader
import xlrd
import hashlib
import os


def handle_uploaded_file(f, name):
    destination = open(os.getcwd() + settings.STATIC_ROOT + name, 'wb+')
    for chunk in f.chunks():
        destination.write(chunk)
    destination.close()


#excel 导入
class ListImportExcelPlugin(BaseAdminPlugin):
    import_excel = False

    def init_request(self, *args, **kwargs):
        return bool(self.import_excel)

    def block_top_toolbar(self, context, nodes):
        nodes.append(loader.render_to_string('xadmin/views/model_list.top_toolbar.import.html', context_instance=context))

    def post(self, __, request, *args, **kwargs):
        if request.FILES.get('excel'):
            kind = request.FILES['excel'].name.split('.')[-1]
            h = hashlib.md5()
            h.update(request.FILES['excel'].name)
            file_name = h.hexdigest() + '.' + kind
            handle_uploaded_file(request.FILES['excel'], file_name)
            #解析文件
            request.file_name = os.getcwd() + settings.STATIC_ROOT + file_name
#             data = xlrd.open_workbook(os.getcwd() + settings.STATIC_ROOT + file_name)
#             table = data.sheets()[0]
#             request.table = table
#             os.remove(os.getcwd() + settings.STATIC_ROOT + file_name)
        return __()


xadmin.site.register_plugin(ListImportExcelPlugin, ListAdminView)


# 1.先安装 xlrd 包
# 2.实例代码
# 
# class ProductAdmin(object):
#     import_excel = True
# 
#     def post(self,request, *args, **kwargs):
#         rp = super(ProductAdmin,self).post(request,args,kwargs)
#         if hasattr(request,"table"):
#             for i in range(request.table.nrows-1):
#                 db = request.table.row_values(i+1)
#                 #db 即每一行excel 表的数据。从第EXCEL二行开始计算。
#                 print db[0]
#         return rp
# xadmin.site.register(Product, ProductAdmin)