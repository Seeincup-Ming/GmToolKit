# -*- coding: utf-8 -*-


import xlrd
import xlwt


excelfilename = "gmcommand.xls"
excelsheetnum = 3

TDatafilename = "T_Data.xls"


class ExcelReader():
    def __init__(self):

        self.excelfilename = excelfilename
        # self.excelsheetnames = excelsheetnames
        self.loadExcel(self.excelfilename)
        pass

    def loadExcel(self, filename):
        try:
            self.excelFile = xlrd.open_workbook(filename)
            self.excelSheets = self.excelFile.sheet_names()
        except Exception, e:
            print str(e)

    def getSheet(self, sheet):
        try:
            if isinstance(sheet, str):
                # print "read sheet by utf-8 str"
                return self.excelFile.sheet_by_name(sheet.decode("utf"))
            elif isinstance(sheet, unicode):
                # print "read sheet by unicode str"
                return self.excelFile.sheet_by_name(sheet)

            return self.excelFile.sheet_by_index(sheet)
        except:
            return None

    def excelData(self):
        # use this funciton to get the dict which is containning the gmcommands paramaters..

        dicttoreturn = {}

        for i in self.excelSheets:
            if i is not None:
                # print isinstance(i, unicode)
                list = {}
                current_sheet = self.getSheet(i)
                # print current_sheet,"test"
                colnames = current_sheet.row_values(0)  # get the colnames
                # print colnames,"test2"
                #ncols = current_sheet.ncols  # lie
                nrows = current_sheet.nrows  # hang

                for rownum in range(1, nrows):
                    row = current_sheet.row_values(rownum)
                    if row:
                        command_info = {}
                        for j in range(len(colnames)):
                            command_info[colnames[j]] = row[j]
                        list[int(row[0])] = command_info

                #print list
                dicttoreturn[i] = list

        return dicttoreturn

class TDataReader():
    def __init__(self):

        self.excelfilename = TDatafilename
        self.loadExcel(self.excelfilename)

    def loadExcel(self,filename):
        try:
            self.excelFile = xlrd.open_workbook(filename)
            self.excelSheets = self.excelFile.sheet_names()
        except Exception, e:
            print str(e)

    def getSheet(self, sheet):
        try:
            if isinstance(sheet, str):
                # print "read sheet by utf-8 str"
                return self.excelFile.sheet_by_name(sheet.decode("utf-8"))
            elif isinstance(sheet, unicode):
                # print "read sheet by unicode str"
                return self.excelFile.sheet_by_name(sheet)

            return self.excelFile.sheet_by_index(sheet)
        except:
            return None
    def getTData(self):
        dicttoreturn = {}
        for i in self.excelSheets:
            if i is not None:
                # print isinstance(i, unicode)
                list = {}
                current_sheet = self.getSheet(i)
                # print current_sheet,"test"
                colnames = current_sheet.row_values(0)  # get the colnames
                # print colnames,"test2"
                #ncols = current_sheet.ncols  # lie
                nrows = current_sheet.nrows  # hang


                for rownum in range(1, nrows):
                    row = current_sheet.row_values(rownum)
                    if row:
                        command_info = {}
                        for j in range(len(colnames)):
                            command_info[colnames[j]] = row[j]
                        list[int(row[0])] = command_info


                dicttoreturn[i] = list

        return dicttoreturn
    def writeTData_group(self,setid,content):
        dictread = self.getTData()
        #    {u'commands': {1: {u'Content': u'$mhp', u'Id': u'1', u'Set_id': u'1'}, 2: {u'Content': u'/cd', u'Id': u'2', u'Set_id': u'1'}, 3: {u'Content': u'$mhp', u'Id': u'3', u'Set_id': u'2'}, 4: {u'Content': u'/cd', u'Id': u'4', u'Set_id': u'2'}, 5: {u'Content': u'$mhp', u'Id': u'5', u'Set_id': u'2'}}, u'group': {1: {u'Content': u'\u7528\u6765\u6d4b\u8bd5\u7ec3\u4e60\u573a\u7684\u6307\u4ee4\u96c6', u'Set_id': u'1'}, 2: {u'Content': u'\u7528\u6765\u6d4b\u8bd55v5\u7684\u6307\u4ee4\u96c6', u'Set_id': u'2'}}}

        tempdict= {}

        newsetid = setid
        while(setid in dictread[u'group'].keys()):
            newsetid = setid + 1

        tempdict[u'Set_id'] = newsetid
        tempdict[u'Content'] = content


        dictread[u'group'][newsetid] = tempdict
        #u'group': {1: {u'Content': u'\u7528\u6765\u6d4b\u8bd5\u7ec3\u4e60\u573a\u7684\u6307\u4ee4\u96c6', u'Set_id': u'1'}, 2: {u'Content': u'\u7528\u6765\u6d4b\u8bd55v5\u7684\u6307\u4ee4\u96c6', u'Set_id': u'2'}
        #print dictread

        keys = dictread.keys()

        wbk = xlwt.Workbook("utf-8")

        groupsheet = wbk.add_sheet(u'group',cell_overwrite_ok=True)

        groupdict = dictread.get(u'group')
        #print len(groupdict),"test2",
        #print groupdict

        groupsheet.write(0,0,u'Set_id')
        groupsheet.write(0,1,u'Content')

        for i,key in enumerate(groupdict):
            #print i ,key,groupdict.get(key)
            linetemp = groupdict.get(key)
            groupsheet.write(i+1,0,unicode(linetemp.get(u'Set_id')))
            groupsheet.write(i+1,1,unicode(linetemp.get(u'Content')))

        ##############添加存储commands的代码
        commandssheet = wbk.add_sheet(u'commands',cell_overwrite_ok=True)

        commandsdict = dictread.get(u'commands')

        commandssheet.write(0,0,u'Id')
        commandssheet.write(0,1,u'Set_id')
        commandssheet.write(0,2,u'Content')

        for i,key in enumerate(commandsdict):
            #print i ,key,commandsdict.get(key)
            linetemp = commandsdict.get(key)
            commandssheet.write(i+1,0,unicode(linetemp.get(u'Id')))
            commandssheet.write(i+1,1,unicode(linetemp.get(u'Set_id')))
            commandssheet.write(i+1,2,unicode(linetemp.get(u'Content')))


        try:
            wbk.save(self.excelfilename)
        except Exception,e:
            print str(e)
    def deleteTData_group(self,setid):

        #{u'commands': {1: {u'Content': u'$mhp', u'Id': u'1', u'Set_id': u'1'},
        # 2: {u'Content': u'/cd', u'Id': u'2', u'Set_id': u'1'},}

        dictread = self.getTData()
        commands = dictread.get(u"commands")
        #print "1",commands
        #newdict = {}
        for key in commands.keys():
            #print "*"
            if commands.get(key).get(u"Set_id") == unicode(setid):
                dictread[u'commands'].pop(key)
        counter = 1
        commands = dictread.get(u"commands")
        #dictlength = len(commands)

        newcommands = {}

        for key in commands.keys():

            tempdict = {}
            tempdict[u"Content"] = commands.get(key).get(u"Content")
            tempdict[u"Id"] = unicode(counter)
            tempdict[u"Set_id"] = commands.get(key).get(u"Set_id")
            newcommands[counter] = tempdict
            counter = counter + 1

        newdictread = {}
        newdictread[u"commands"] = newcommands


        groups = dictread.get(u'group')
        #print "*",groups

        if groups.has_key(setid):
            dictread.get(u'group').pop(setid)

        counter = 1
        newgroups = dictread.get(u'group')
        groups_temp ={}
        for key in newgroups.keys():
            tempdict2 = {}
            tempdict2[u"Content"] = newgroups.get(key).get(u"Content")
            tempdict2[u"Set_id"] = unicode(counter)
            groups_temp[counter] = tempdict2
            counter = counter + 1

        newdictread[u"group"] = groups_temp


        wbk = xlwt.Workbook("utf-8")

        commandssheet = wbk.add_sheet(u'commands',cell_overwrite_ok=True)

        commandsdict = newdictread.get(u'commands')

        commandssheet.write(0,0,u'Id')
        commandssheet.write(0,1,u'Set_id')
        commandssheet.write(0,2,u'Content')

        for i,key in enumerate(commandsdict):

            linetemp = commandsdict.get(key)
            commandssheet.write(i+1,0,unicode(linetemp.get(u'Id')))
            commandssheet.write(i+1,1,unicode(linetemp.get(u'Set_id')))
            commandssheet.write(i+1,2,unicode(linetemp.get(u'Content')))
        #######添加存储group的代码

        groupsheet = wbk.add_sheet(u'group',cell_overwrite_ok=True)

        groupdict = newdictread.get(u'group')

        groupsheet.write(0,0,u'Set_id')
        groupsheet.write(0,1,u'Content')

        for i,key in enumerate(groupdict):
            #print i ,key,groupdict.get(key)
            linetemp = groupdict.get(key)
            groupsheet.write(i+1,0,unicode(linetemp.get(u'Set_id')))
            groupsheet.write(i+1,1,unicode(linetemp.get(u'Content')))

        try:
            wbk.save(self.excelfilename)
        except Exception,e:
            print str(e)

        pass

    def deleteTData_command(self,leftid,rightid):

        dictread = self.getTData()
        commands = dictread.get(u"commands")
        #print commands

        counter = rightid
        commandid = None

        for key in commands.keys():
            oneline = commands.get(key)
            if oneline.get(u"Set_id") == unicode(str(leftid)):
                counter = counter - 1
                if counter == 0:
                    commandid = int(oneline.get(u"Id"))

        #print commandid,"here"

        newdict = {}

        if commands.has_key(commandid):
            dictread[u'commands'].pop(commandid)
            commands = dictread.get(u"commands")
            dictkeys = dictread[u'commands'].keys()
            dictkeys.sort()


            for key in dictkeys:
                if key > commandid:
                    oneline = commands.get(key)
                    temponeline = {}
                    temponeline[u'Content'] = oneline.get(u'Content')
                    temponeline[u'Id'] = unicode(int(oneline.get(u'Id'))-1)
                    temponeline[u'Set_id'] = oneline.get(u'Set_id')
                    newdict[key-1] = temponeline
                else:
                    oneline = commands.get(key)
                    newdict[key] = oneline

        newdictread = {}
        newdictread[u"commands"] = newdict

        # 下面是将dictread写会文件的代码。
        wbk = xlwt.Workbook("utf-8")

        commandssheet = wbk.add_sheet(u'commands',cell_overwrite_ok=True)

        commandsdict = newdictread.get(u'commands')

        commandssheet.write(0,0,u'Id')
        commandssheet.write(0,1,u'Set_id')
        commandssheet.write(0,2,u'Content')

        for i,key in enumerate(commandsdict):
            #print i ,key,commandsdict.get(key)
            linetemp = commandsdict.get(key)
            commandssheet.write(i+1,0,unicode(linetemp.get(u'Id')))
            commandssheet.write(i+1,1,unicode(linetemp.get(u'Set_id')))
            commandssheet.write(i+1,2,unicode(linetemp.get(u'Content')))

        #######添加存储group的代码

        groupsheet = wbk.add_sheet(u'group',cell_overwrite_ok=True)

        groupdict = dictread.get(u'group')
        #print len(groupdict),"test2",
        #print groupdict

        groupsheet.write(0,0,u'Set_id')
        groupsheet.write(0,1,u'Content')

        for i,key in enumerate(groupdict):
            #print i ,key,groupdict.get(key)
            linetemp = groupdict.get(key)
            groupsheet.write(i+1,0,unicode(linetemp.get(u'Set_id')))
            groupsheet.write(i+1,1,unicode(linetemp.get(u'Content')))

        try:
            wbk.save(self.excelfilename)
        except Exception,e:
            print str(e)

        pass

    def writeTData_command(self,setid,content):
        dictread = self.getTData()
        #    {u'commands': {1: {u'Content': u'$mhp', u'Id': u'1', u'Set_id': u'1'},
        # 2: {u'Content': u'/cd', u'Id': u'2', u'Set_id': u'1'},
        # 3: {u'Content': u'$mhp', u'Id': u'3', u'Set_id': u'2'},
        # 4: {u'Content': u'/cd', u'Id': u'4', u'Set_id': u'2'},
        # 5: {u'Content': u'$mhp', u'Id': u'5', u'Set_id': u'2'}},
        # u'group': {1: {u'Content': u'\u7528\u6765\u6d4b\u8bd5\u7ec3\u4e60\u573a\u7684\u6307\u4ee4\u96c6', u'Set_id': u'1'}, 2: {u'Content': u'\u7528\u6765\u6d4b\u8bd55v5\u7684\u6307\u4ee4\u96c6', u'Set_id': u'2'}}}

        tempdict= {}

        id = len(dictread.get(u"commands"))+1

        tempdict[u'Id'] = unicode(id)
        tempdict[u'Set_id'] = unicode(setid)
        tempdict[u'Content'] = content

        dictread[u'commands'][id] = tempdict



        # 下面是将dictread写会文件的代码。
        wbk = xlwt.Workbook("utf-8")

        commandssheet = wbk.add_sheet(u'commands',cell_overwrite_ok=True)

        commandsdict = dictread.get(u'commands')

        commandssheet.write(0,0,u'Id')
        commandssheet.write(0,1,u'Set_id')
        commandssheet.write(0,2,u'Content')

        for i,key in enumerate(commandsdict):
            #print i ,key,commandsdict.get(key)
            linetemp = commandsdict.get(key)
            commandssheet.write(i+1,0,unicode(linetemp.get(u'Id')))
            commandssheet.write(i+1,1,unicode(linetemp.get(u'Set_id')))
            commandssheet.write(i+1,2,unicode(linetemp.get(u'Content')))

        #######添加存储group的代码

        groupsheet = wbk.add_sheet(u'group',cell_overwrite_ok=True)

        groupdict = dictread.get(u'group')
        #print len(groupdict),"test2",
        #print groupdict

        groupsheet.write(0,0,u'Set_id')
        groupsheet.write(0,1,u'Content')

        for i,key in enumerate(groupdict):
            #print i ,key,groupdict.get(key)
            linetemp = groupdict.get(key)
            groupsheet.write(i+1,0,unicode(linetemp.get(u'Set_id')))
            groupsheet.write(i+1,1,unicode(linetemp.get(u'Content')))

        try:
            wbk.save(self.excelfilename)
        except Exception,e:
            print str(e)




if __name__ == '__main__':
    # print monster_sequence
    er = TDataReader()
    #er.deleteTData_group(3)  #删除某个setid的commands
    #er.writeTData_group(5,u"$$$$$$$$$$$$$$$$$$$$$$")
    er.writeTData_command(4,"$sadfas")
    #er.deleteTData_command(3) # 删除某个id的例子。一次只能一个。
    #er.deleteTData_command(3)







