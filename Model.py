
import xmlrpc.client
import datetime
from tkinter import messagebox


class Connection:

    def __init__(self):

        self.testServer = ["gdfusion-tests-1-433058", "https://gdfusion-tests-1-433058.dev.odoo.com"]
        self.productionServer = ["gdfusion-stage-260946", "https://gdfusion.odoo.com"]

        self.USER = "mfabela@gdfusion.com.mx"  # email address
        self.PASS = "admin"  # password
        self.PORT = "443"  # port

        result = messagebox.askyesno(message="¿Use test server?", title="Server Selection")

        if result:
            # Test Server
            self.DATA = self.testServer[0]  # db name
            self.URL = self.testServer[1]  # base url
        else:
            # Production Server
            self.DATA = self.productionServer[0]  # db name
            self.URL = self.productionServer[1]  # base url

        self.URL_COMMON = "{}:{}/xmlrpc/2/common".format(
            self.URL, self.PORT)
        self.URL_OBJECT = "{}:{}/xmlrpc/2/object".format(
            self.URL, self.PORT)


    def startConnection(self):
        self.ODOO_COMMON = xmlrpc.client.ServerProxy(self.URL_COMMON)
        self.ODOO_OBJECT = xmlrpc.client.ServerProxy(self.URL_OBJECT)
        self.UID = self.ODOO_COMMON.authenticate(
            self.DATA
            , self.USER
            , self.PASS
            , {})
        messagebox.showinfo("Connection", "User: " + self.USER + " connected to: " + self.URL)



