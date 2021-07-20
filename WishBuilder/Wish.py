#!/usr/bin/env python
# -*- coding: utf-8 -*-
from bottle import *
from datetime import datetime
from xlrd import xldate_as_tuple
import xlrd, os, requests, json
import WishBuilder


HTML = """
<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>愿望城堡积分添加</title>
        <style type="text/css">
            body{
            font-size:14px;
            align-text:center;
            background-image: url("https://cn.bing.com/th?id=OHR.LakePinatubo_ZH-CN5947011761_1920x1080.jpg&rf=LaDigue_1920x1080.jpg&pid=hp");
            background-repeat:no-repeat;
            background-attachment: fixed;
            background-size: cover;
            opacity: 77%;
            }
            input{ 
            vertical-align:middle;
            margin:0;
            padding:0
            }
            .file-box{
            position:relative;
            width:340px;
            margin:0px auto;
            }
            .Boxposition{
            position:absolute;
            top:35%;
            left:37%;
            }
            .txt{
            height:22px;
            border:1px solid #cdcdcd;
            width:180px;
            }
            .btn{
            background-color:#FFF;
            border:1px solid #CDCDCD;
            height:24px;
            width:70px;
            }
            .file{
            position:absolute;
            top:0;
            right:80px;
            height:24px;
            filter:alpha(opacity:0);
            opacity: 0;
            width:260px
            }
        </style>
    </head>
<body>
    <div class="Boxposition">
        <div class="file-box">
            <form action="/upload" method="post" enctype="multipart/form-data">
                <input type='text' name='textfield' id='textfield' class='txt' />
                <input type='button' class='btn' value='浏览...' />
                <input type="file" name="fileField" class="file" id="fileField" size="28" onchange="document.getElementById('textfield').value=this.value" />
                <input type="submit" name="submit" class="btn" value="确认上传" onclick=""/>
            </form>
        </div>
        
        <div style='position:relative; width:400px; margin-top:50px;margin-left:50px;'>
            <p style="font-size:20px; color:white">
                <b>愿望城堡积分添加 Excel 文件格式如下：</b>
            </p>
            <p style='color:yellow'>
                <b>注：Excel需要去除表头（伙伴工号、愿望积分、备注内容） </b>
            </p>  
        </div>
        <div style='border:1px dashed white; position:relative; width:360px; left:50px; margin-top:25px'>  
            <p style='color : white'> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 第一列  &nbsp;&nbsp;&nbsp; 第二列 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  第三列</p>
            <p style='color : white'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;伙伴工号&nbsp;&nbsp;&nbsp;愿望积分&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;备注内容</p>
            <p style='color : white'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;607577&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;500&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;小额激励-SC金点子</p>
            <p style='color : white'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;617681&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;200&nbsp;&nbsp;&nbsp;&nbsp;小额激励-苏皖区技术PK赛</p>
            <p style='color : white'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;600588&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;500&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;读书会</p>
        </div>
    </div>
</body>
</html>
"""

Success_page="""
<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>愿望城堡积分添加</title>
        
        <style type="text/css">
            body{
            font-size:14px;
            align-text:center;
            background-image: url("https://cn.bing.com/th?id=OHR.AnnularEclipse_ZH-CN2345201060_1920x1080.jpg&rf=LaDigue_1920x1080.jpg&pid=hp");
            background-repeat:no-repeat;
            background-attachment: fixed;
            background-size: cover;
            opacity: 77%;
            }
            .Boxposition{
            position:absolute;
            top:70%;
            left:47%;
            }
            .itsform{
            position:absolute;
            top:90%;
            left:18%;
            }
            input{ 
            vertical-align:middle;
            margin:0;
            padding:0
            }
            .file-box{
            position:relative;
            width:340px;
            margin:0px auto;
            }
            .txt{
            height:22px;
            border:1px solid #cdcdcd;
            width:180px;
            }
            .btn{
            background-color:#FFF;
            border:1px solid #CDCDCD;
            height:24px;
            width:70px;
            }
            .file{
            position:absolute;
            top:0;
            right:80px;
            height:24px;
            filter:alpha(opacity:0);
            opacity: 0;
            width:260px
            }
        </style>
    </head>
<body>

    <div class='Boxposition',style='border:1px dashed transparent; left:44%; width:200px; position:relative; margin-top: 400px; '>  
    <p style='font-size:20px; color:white; '>文件上传成功</p>
        <div class="itsform">
            <form action="/upwish" method="get" enctype="multipart/form-data">
                <input type="submit" onClick="alert('已提交,接口请求中... ... 请勿重复提交！！！')" name="submit" class="btn" value="确认提交" onclick=""/>
            </form>
	    </div>
    </div>
</body>
</html>

"""


fail_page="""
<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>愿望城堡积分添加</title>
        
        <style type="text/css">
            body{
            font-size:14px;
            text-align:center;
            align-text:center;
            background-image: url("https://cn.bing.com/th?id=OHR.AnnularEclipse_ZH-CN2345201060_1920x1080.jpg&rf=LaDigue_1920x1080.jpg&pid=hp");
            background-repeat:no-repeat;
            background-attachment: fixed;
            background-size: cover;
            opacity: 77%;
            }
            input{ 
            vertical-align:middle;
            margin:0;
            padding:0
            }
            .file-box{
            position:relative;
            width:340px;
            margin:0px auto;
            }
            .Boxposition{
            position:absolute;
            top:35%;
            left:34%;
            width:100%;
            height:100%;
            }
            .txt{
            height:22px;
            border:1px solid #cdcdcd;
            width:180px;
            }
            .btn{
            background-color:#FFF;
            border:1px solid #CDCDCD;
            height:24px;
            width:70px;
            }
            .file{
            position:absolute;
            top:0;
            right:80px;
            height:24px;
            filter:alpha(opacity:0);
            opacity: 0;
            width:260px
            }
        </style>
    </head>
<body>
    <div class='Boxposition'>
        <div style='border:1px dashed transparent; width:600px; position:relative; text-align:center'>  
            <p style='font-size:20px; color:white;'>上传文件失败,已存在相同名称文件,请删除后重试！</p>
            <div id="nav1">
                <p style='color:white;'>点击下方按钮删除同名文件并返回重新上传文件</p>
                <form action="/delete" method="get" enctype="multipart/form-data">
                    <input type="submit" name="delete" class="btn" value="确认删除" onclick=""/>
                </form>
            </div>
        </div>
    </div>
</body>
</html>
"""

base_path = os.path.dirname(os.path.realpath(__file__))  # 获取脚本路径

upload_path = os.path.join(base_path, 'upload')  # 上传文件目录
if not os.path.exists(upload_path):
    os.makedirs(upload_path)


@route('/', method='GET')
@route('/upload', method='GET')
@route('/upwish', method='GET')
@route('/delete', method='GET')
@route('/index.html', method='GET')
@route('/upload.html', method='GET')

@route('index.html',method='GET')
def index():
    """显示上传页"""
    return HTML


@route('/upload', method='POST')
def do_upload():
    """处理上传文件"""
    filedata = request.files.get('fileField')

    if filedata.file:
        global file_path,file_name
        file_name = filedata.filename
        file_path = os.path.join(upload_path, filedata.filename)
        print("file_path",file_path)
        try:
            filedata.save(file_path)  # 上传文件写入
        except IOError:
            return fail_page
        return Success_page
    else:
        return fail_page


#调用WishBuilder.py文件执行添加积分
try:
    @route('/upwish', methed='GET')
    def do_UpWish():
        data = xlrd.open_workbook(file_path)
        table = data.sheets()[0]
        tables=[ ]
        result = WishBuilder.import_excel(table=table,tables=tables)
        result_dic = {'result':result}
        finash = """
        <!DOCTYPE html>
        <html>
            <head>
                <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
                <title>愿望城堡积分添加</title>

                <style type="text/css">
                    body{
                    font-size:14px;
                    align-text:center;
                    background-image: url("https://cn.bing.com/th?id=OHR.AnnularEclipse_ZH-CN2345201060_1920x1080.jpg&rf=LaDigue_1920x1080.jpg&pid=hp");
                    background-repeat:no-repeat;
                    background-attachment: fixed;
                    background-size: cover;
                    opacity: 77%;
                    }
                    input{ 
                    vertical-align:middle;
                    margin:0;
                    padding:0
                    }
                    .file-box{
                    position:relative;
                    width:340px;
                    margin:0px auto;
                    }
                    .Boxposition{
                    position:absolute;
                    top:10%;
                    width:100%;
                    height:100%;
                    text-align:center;
                    }
                    .txt{
                    height:22px;
                    border:1px solid #cdcdcd;
                    width:180px;
                    }
                    .btn{
                    background-color:#FFF;
                    border:1px solid #CDCDCD;
                    height:24px;
                    width:70px;
                    }
                    .file{
                    position:absolute;
                    top:0;
                    right:80px;
                    height:24px;
                    filter:alpha(opacity:0);
                    opacity: 0;
                    width:260px
                    }
                </style>
            </head>
        <body>
            <div class='Boxposition'>
                <div>
                <p style='font-size:30px; color:white;'>愿望城堡伙伴积分已添加完毕(请检查 Status_code)</p>
                <marquee scrollamount="10"><span style="font-weight: bolder;font-size: 16px;color: yellow;">注： Status_code ≠ 200 为添加失败,请确认后联系相应人员进行处理添加</span></marquee>
                <div>
                    <div>
                    <p style='color:yellow;'> \
                    <br> \
                    </p>\
                </div>""" + \
                    "<div style='border:1px dashed transparent; width:70%; margin-left:36%; text-align:left;'>" \
                    "<br>" \
                    "<p style='color:white;'>" \
                    "<b> 添加详情如下: </b>" \
                    "<br>" \
                    "</p>" \
                    "<p style='color:white'>{}</p>" \
                    "</div>"\
                        .format(str(result_dic['result']).replace('{','').replace('}','<br><br>').replace('[','').replace(']','').replace(',','').replace('"',' ').replace("'","").replace('"',"").replace('note:',"<br>备注:")) \
                 + """
                <br>
                <form action="/delete" method="get" enctype="multipart/form-data">
                    <input type="submit" name="back"')" class="btn" value="返回首页" onclick=""/>
                </form>
                <br>
                <br>
                <br>
            </div>
        </body>
        </html>
        """
        return finash


except:
    print('初次执行不存在目标Excel文件跳过！！！')

#删除已存在同名文件
try:
    @route('/delete', methed='GET')
    def do_delete():
        os.remove(os.path.join(upload_path, file_name))
        return HTML
except:
    print('初次执行不存在目标Excel文件跳过！！！')



@route('/favicon.ico', method='GET')
def server_static():
    """处理网站图标文件, 找个图标文件放在脚本目录里"""
    return static_file('favicon.ico', root=base_path)



@error(404)
def error404(error):
    """处理错误信息"""
    return '404 发生页面错误, 未找到内容'

run(host='10.20.221.63',port=8080)
#run(port=8080, reloader=False)  # reloader设置为True可以在更新代码时自动重载
