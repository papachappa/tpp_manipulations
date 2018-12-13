import wmi


class RemoveApp():
    def __init__(self, name):
        self.w = wmi.WMI()
        self.name = name

    def find_app(self):
        for app in self.w.Win32_Product():
            if app.Name.startswith(self.name):
                return app.Name

    def uninstall(self, app):
        for product in self.w.Win32_Product(Name=app):
           print("Uninstalling " + app + "...")
           product.Uninstall()
           print("The app uninstalled")

if __name__ == "__main__":
    name = RemoveApp("Venafi User Portal")
    n = name.find_app()
    name.uninstall(n)
    name = RemoveApp("Venafi Trust Protection Platform")
    n = name.find_app()
    name.uninstall(n)
