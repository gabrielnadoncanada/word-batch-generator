import win32com.client as win32
from win32com.client import constants
outlook = win32.gencache.EnsureDispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")
ns.Logon("", "", False, False)
print("[OK] Session MAPI active")

mail = outlook.CreateItem(constants.olMailItem)
mail.To = "gabrielnadoncanada@gmail.com"
mail.Subject = "Test COM"
mail.HTMLBody = "Hello via COM"
mail.Save()
print("[OK] Draft créé")
# mail.Send()  # décommente pour tester l'envoi
