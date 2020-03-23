import win32com.client
from win32com import client

#C:\Users\MYost\Desktop\Diff\AcceptedRedlines
path = r'C:\Users\MYost\Desktop\Diff\Quizzes\\'
# note the \\ at the end of the path name to prevent a SyntaxError

#Create the Application word
Application=win32com.client.gencache.EnsureDispatch("Word.Application")



# Compare documents
Application.CompareDocuments(Application.Documents.Open(path + "V0\Group13_Master_V0.doc"),
                             Application.Documents.Open(path + "V1\Group13_Master_V1.doc"))

Application.ActiveDocument.ActiveWindow.View.Type = 3
# Save the comparison document as "Comparison.docx"
Application.ActiveDocument.SaveAs (FileName = path + "Comparison.docx")
# Don't forget to quit your Application
Application.Quit()
