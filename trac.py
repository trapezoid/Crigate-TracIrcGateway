import clr

import System
import Misuzilla.Crigate
import Misuzilla.Crigate.Scripting

from System                     import *
from System.Collections.Generic import *
from System.Diagnostics         import Trace

clr.AddReference("CookComputing.XmlRpcV2.dll")
clr.AddReference("BTS.Trac.dll")

from BTS.Trac import *

from Misuzilla.Crigate               import ConsoleHandler, Console, Context
from Misuzilla.Crigate.Configuration import IConfiguration, ConfigurationPropertyInfo
from Misuzilla.Crigate.Scripting     import DLRIntegrationHandler, DLRBasicConfiguration, DLRContextHelper

class TracContext(Context):
    def Initialize(self):
        self.config = TracContext.GetConfig()

    @staticmethod
    def GetConfig():
        return DLRBasicConfiguration(CurrentSession, "TracContext", Array[ConfigurationPropertyInfo](
            [
                ConfigurationPropertyInfo("TracUrl", "TracのURL", System.String, True, None),
                ConfigurationPropertyInfo("TracUsername", "Tracのログインユーザー", System.String, True, None),
                ConfigurationPropertyInfo("TracPassword", "Tracのログインパスワード", System.String, True, None),
            ]))

    @staticmethod
    def Connect():
        config = TracContext.GetConfig()
        Trac.Connect(config.GetValue("TracUrl"), config.GetValue("TracUsername"), config.GetValue("TracPassword"))

    def GetCommands(self):
        dict = Context.GetCommands(self)
        dict["Component"] = "コンポーネントコンテキストに切り替えます"
        dict["Ticket"] = "チケットコンテキストに切り替えます"
        dict["Milestone"] = "マイルストーンコンテキストに切り替えます"
        return dict

    def OnUninitialize(self):
        pass

    def get_Configurations(self):
        return Array[IConfiguration]([ self.config ])

    # Implementation
    def Ticket(self, args):
        nextContext = self.Console.GetContext(DLRContextHelper.Wrap(CurrentSession, "Ticket", TicketContext), CurrentSession)
        self.Console.PushContext(nextContext)

    def Component(self, args):
        nextContext = self.Console.GetContext(DLRContextHelper.Wrap(CurrentSession, "Component", ComponentContext), CurrentSession)
        self.Console.PushContext(nextContext)

    def Milestone(self, args):
        nextContext = self.Console.GetContext(DLRContextHelper.Wrap(CurrentSession, "Milestone", MilestoneContext), CurrentSession)
        self.Console.PushContext(nextContext)

class ComponentContext(Context):
    def Initialize(self):
        self.rootConfig = TracContext.GetConfig()
        self.config = ComponentContext.GetConfig()

    @staticmethod
    def GetConfig():
        return DLRBasicConfiguration(CurrentSession, "ComponentContext", Array[ConfigurationPropertyInfo](
            [
                ConfigurationPropertyInfo("DefaultComponent", "コンポーネントの初期値", System.String, None, None),
                ConfigurationPropertyInfo("DefaultMilestone", "マイルストーンの初期値", System.String, None, None),
                ConfigurationPropertyInfo("DefaultPriority", "優先度の初期値", System.String, None, None),
            ]))

    def GetCommands(self):
        dict = Context.GetCommands(self)
        return dict

    def OnUninitialize(self):
        pass

    def get_Configurations(self):
        return Array[IConfiguration]([ self.config ])

    def Query(self, args):
        pass

    def List(self, args):
        TracContext.Connect()
        for component in Component().GetAll() :
            self.Console.NotifyMessage(("Component: %s") % (component))

    def Edit(self, args):
        pass

    def New(self, args):
        TracContext.Connect()
        nextContext = self.Console.GetContext(DLRContextHelper.Wrap(CurrentSession, "NewComponent", CreateTicketContext), CurrentSession)
        self.Console.PushContext(nextContext)

class MilestoneContext(Context):
    def Initialize(self):
        self.rootConfig = TracContext.GetConfig()
        self.config = MilestoneContext.GetConfig()

    @staticmethod
    def GetConfig():
        return DLRBasicConfiguration(CurrentSession, "MilestoneContext", Array[ConfigurationPropertyInfo](
            [
            ]))

    def GetCommands(self):
        dict = Context.GetCommands(self)
        dict["New"] = "マイルストーンを登録します"
        dict["List"] = "マイルストーン一覧を表示します"
        dict["Edit"] = "マイルストーンを編集します"
        dict["Get"] = "マイルストーンを閲覧します"
        dict["Delete"] = "マイルストーンを削除します"
        return dict

    def OnUninitialize(self):
        pass

    def get_Configurations(self):
        return Array[IConfiguration]([ self.config ])

    def Get(self, args):
        TracContext.Connect()

        m = Milestone()
        m.Get(args)
        self.Console.NotifyMessage((u"%s") % (m.Name))
        self.Console.NotifyMessage((u"    Description : %s") % (m.Description))
        self.Console.NotifyMessage((u"    Due         : %s") % (m.Due))
        self.Console.NotifyMessage((u"    Completed   : %s") % (m.Conmleted))

    def List(self, args):
        TracContext.Connect()
        for milestone in Milestone().GetAll() :
            self.Console.NotifyMessage(("Milestone: %s") % (milestone))

    def Edit(self, args):
        TracContext.Connect()
        nextContext = self.Console.GetContext(DLRContextHelper.Wrap(CurrentSession, "EditMilestone", CreateMilestoneContext), CurrentSession)
        self.Console.PushContext(nextContext)
        loadMethod = nextContext.GetCommand("Load")
        loadMethod.Invoke(nextContext, tuple([args]))

    def Delete(self, args):
        TracContext.Connect()
        Milestone.Delete(args)
        self.Console.NotifyMessage(("Milestone: %s を削除しました") % (args))

    def New(self, args):
        TracContext.Connect()
        nextContext = self.Console.GetContext(DLRContextHelper.Wrap(CurrentSession, "NewMilestone", CreateMilestoneContext), CurrentSession)
        self.Console.PushContext(nextContext)

class CreateMilestoneContext(Context):
    def Initialize(self):
        self.rootConfig = TracContext.GetConfig()
        self.milestoneConfig = MilestoneContext.GetConfig()

        self.config = DLRBasicConfiguration(self.CurrentSession, "NewMilestone", Array[ConfigurationPropertyInfo](
            [
                ConfigurationPropertyInfo("Name", "マイルストーン名", System.String, None, None),
                ConfigurationPropertyInfo("Description", "説明", System.String, None, None),
                ConfigurationPropertyInfo("Due", "完了期限", System.String, None, None),
                ConfigurationPropertyInfo("Completed", "完了日", System.String, None, None),
            ]))
        self.milestone = Milestone()
        self.isEdit = False

    def GetCommands(self):
        dict = Context.GetCommands(self)
        dict["Save"] = ("マイルストーンを%sします") % ("更新" if self.isEdit else "登録")
        return dict

    def OnUninitialize(self):
        self.config.SetValue("Name","")
        self.config.SetValue("Description","")
        self.config.SetValue("Due","")
        self.config.SetValue("Completed","")

    def get_Configurations(self):
        return Array[IConfiguration]([ self.config ])

    def __isUnixTimestamp(self, time):
        if(time.CompareTo(DateTime(1970, 1, 1)) > 0):
            return True
        else:
            return False

    def Load(self, args):
        self.milestone = Milestone()
        self.milestone.Get(args)
        self.config.SetValue("Name", self.milestone.Name)
        self.config.SetValue("Description", self.milestone.Description)

        if(self.__isUnixTimestamp(self.milestone.Due)):
            self.config.SetValue("Due", self.milestone.Due)
        else:
            self.config.SetValue("Due", None)

        if(self.__isUnixTimestamp(self.milestone.Conmleted)):
            self.config.SetValue("Completed", self.milestone.Conmleted)
        else:
            self.config.SetValue("Completed", None)

        self.isEdit = True

    def Save(self, args):
        self.milestone.Name = self.config.GetValue("Name")
        self.milestone.Description = self.config.GetValue("Description")

        if(self.config.GetValue("Due") != None):
            self.milestone.Due = DateTime.Parse(self.config.GetValue("Due"))
            if(self.__isUnixTimestamp(self.milestone.Due) == False):
                self.Console.NotifyMessage("日付の指定が不正です")
                self.Exit()

        if(self.config.GetValue("Completed") != None):
            self.milestone.Conmleted = DateTime.Parse(self.config.GetValue("Completed"))
            if(self.__isUnixTimestamp(self.milestone.Conmleted) == False):
                self.Console.NotifyMessage("日付の指定が不正です")
                self.Exit()

        if(self.isEdit):
            self.milestone.Update()
        else:
            self.milestone.Create()

        self.Console.NotifyMessage(("マイルストーン: %s を%sしました") % (self.milestone.Name, "更新" if self.isEdit else "登録"))
        self.Exit()

class TicketContext(Context):
    def Initialize(self):
        self.rootConfig = TracContext.GetConfig()
        self.config = TicketContext.GetConfig()

    @staticmethod
    def GetConfig():
        return DLRBasicConfiguration(CurrentSession, "TicketContext", Array[ConfigurationPropertyInfo](
            [
                ConfigurationPropertyInfo("DefaultComponent", "コンポーネントの初期値", System.String, None, None),
                ConfigurationPropertyInfo("DefaultMilestone", "マイルストーンの初期値", System.String, None, None),
                ConfigurationPropertyInfo("DefaultPriority", "優先度の初期値", System.String, None, None),
                ConfigurationPropertyInfo("DefaultResolution", "クローズ時の解決方法の初期値", System.String, None, None),
            ]))

    def GetCommands(self):
        dict = Context.GetCommands(self)
        dict["New"] = "チケットを登録します"
        dict["List"] = "未完了のチケット一覧を表示します"
        dict["Edit"] = "チケットを編集します"
        dict["Get"] = "チケットを閲覧します"
        dict["Delete"] = "チケットを削除します"
        dict["Query"] = "指定した条件にマッチするチケット一覧を表示します"
        return dict

    def OnUninitialize(self):
        pass

    def get_Configurations(self):
        return Array[IConfiguration]([ self.config ])

    def Query(self, args):
        TracContext.Connect()

        ticketList = Ticket.Query(args)
        for id in ticketList:
            t = Ticket()
            t.Get(id)
            self.Console.NotifyMessage((u"#%d %s") % (t.ID, t.Summary))

    def List(self, args):
        self.Query("status!=closed")

    def Delete(self, args):
        TracContext.Connect()
        t = Ticket()
        t.Get(int(args))
        self.Console.NotifyMessage((u"#%d %s を削除しました") % (t.ID, t.Summary))
        t.Delete()

    def Edit(self, args):
        TracContext.Connect()
        nextContext = self.Console.GetContext(DLRContextHelper.Wrap(CurrentSession, "EditTicket", CreateTicketContext), CurrentSession)
        self.Console.PushContext(nextContext)
        loadMethod = nextContext.GetCommand("Load")
        loadMethod.Invoke(nextContext, tuple([args]))

    def Get(self, args):
        TracContext.Connect()

        t = Ticket()
        t.Get(int(args))
        self.Console.NotifyMessage((u"#%d %s") % (t.ID, t.Summary))
        self.Console.NotifyMessage((u"    Owner    : %s") % (t.Owner))
        self.Console.NotifyMessage((u"    Type     : %s") % (t.Type))
        self.Console.NotifyMessage((u"    Priority : %s") % (t.Priority))
        self.Console.NotifyMessage((u"    Milestone: %s") % (t.Milestone))
        self.Console.NotifyMessage((u"    Component: %s") % (t.Component))
        self.Console.NotifyMessage((u"    Status   : %s") % (t.Status))

    def New(self, args):
        TracContext.Connect()

        nextContext = self.Console.GetContext(DLRContextHelper.Wrap(CurrentSession, "NewTicket", CreateTicketContext), CurrentSession)
        self.Console.PushContext(nextContext)

class CreateTicketContext(Context):
    def Initialize(self):
        self.rootConfig = TracContext.GetConfig()
        self.ticketConfig = TicketContext.GetConfig()

        self.config = DLRBasicConfiguration(self.CurrentSession, "NewTicket", Array[ConfigurationPropertyInfo](
            [
                ConfigurationPropertyInfo("Summary", "概要", System.String, None, None),
                ConfigurationPropertyInfo("Description", "説明", System.String, "", None),
                ConfigurationPropertyInfo("Priority", "優先度", System.String, self.ticketConfig.GetValue("DefaultPriority"), None),
                ConfigurationPropertyInfo("Milestone", "マイルストーン", System.String, self.ticketConfig.GetValue("DefaultMilestone"), None),
                ConfigurationPropertyInfo("Component", "コンポーネント", System.String, self.ticketConfig.GetValue("DefaultComponent"), None),
            ]))
        self.editConfig = DLRBasicConfiguration(self.CurrentSession, "EditTicket", Array[ConfigurationPropertyInfo](
            [
                ConfigurationPropertyInfo("Resolution", "解決方法", System.String, self.rootConfig.GetValue("DefaultResolution"), None),
                ConfigurationPropertyInfo("Status", "状態", System.String, None, None),
            ]))

        self.ticket = Ticket()
        self.isEdit = False

    def GetCommands(self):
        dict = Context.GetCommands(self)
        if(self.isEdit):
            dict["Close"] = "指定されたアクションでチケットをクローズします"
            dict["Save"] = "指定したコメントでチケットを更新します"
        else:
            dict["Save"] = "指定したコメントでチケットを登録します"

        return dict

    def OnUninitialize(self):
        self.config.SetValue("Summary",None)
        self.config.SetValue("Description",None)
        self.config.SetValue("Priority", ticketConfig.GetValue("DefaultPriority"))
        self.config.SetValue("Milestone", ticketConfig.GetValue("DefaultMilestone"))
        self.config.SetValue("Component", ticketConfig.GetValue("DefaultComponent"))

    def get_Configurations(self):
        if(self.isEdit):
            return Array[IConfiguration]([ self.config, self.editConfig ])
        else:
            return Array[IConfiguration]([ self.config ])

    def Accept(self, comment = ""):
        if(self.isEdit == False):
            return
        self.ticket.Status = "accepted"
        self.Save(comment)

    def Close(self, comment = ""):
        if(self.isEdit == False):
            return
        if(self.editConfig.GetValue("Resolution") == None):
            self.Console.NotifyMessage("Resolutionが未指定です")
            return
        self.ticket.Resolution = self.editConfig.GetValue("Resolution")
        self.ticket.Status = "closed"
        self.Save(comment)

    def Load(self, args):
        self.ticket.Get(int(args))
        self.config.SetValue("Summary", self.ticket.Summary)
        self.config.SetValue("Description", self.ticket.Description)
        self.config.SetValue("Priority", self.ticket.Priority)
        self.config.SetValue("Milestone", self.ticket.Milestone)
        self.config.SetValue("Component", self.ticket.Component)
        self.editConfig.SetValue("Status", self.ticket.Status)
        self.editConfig.SetValue("Resolution", self.ticket.Resolution)

        self.isEdit = True

    def Save(self, comment = ""):
        if(self.editConfig.GetValue("Summary") == None):
            self.Console.NotifyMessage("Summaryが未指定です")
            return

        self.ticket.Summary = self.config.GetValue("Summary")
        self.ticket.Description = self.config.GetValue("Description")
        self.ticket.Priority = self.config.GetValue("Priority")
        self.ticket.Milestone = self.config.GetValue("Milestone")
        self.ticket.Component = self.config.GetValue("Component")

        if(self.isEdit):
            self.ticket.Update(comment)
        else:
            self.ticket.Create()

        self.Console.NotifyMessage(("Ticket: #%dを%sしました") % (self.ticket.ID, "更新" if self.isEdit else "登録"))
        self.Exit()

# コンソールチャンネルを追加する
console = Misuzilla.Crigate.Console()
console.Initialize(CurrentSession)
console.Attach("#Trac", DLRContextHelper.Wrap(CurrentSession, "Trac", TracContext))
# 後片付け
CurrentSession.HandlerLoader.GetHandler[DLRIntegrationHandler]().BeforeUnload += lambda sender, e: console.Detach()
