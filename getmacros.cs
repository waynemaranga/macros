using Microsoft.Office.Interop.Word;
using Microsoft.Vbe.Interop;

...

public List<string> GetMacrosFromDoc()
{
    Document doc = GetWordDoc(@"C:\Temp\test.docm");

    List<string> macros = new List<string>();

    VBProject prj;
    CodeModule code;
    string composedFile;

    prj = doc.VBProject;
    foreach (VBComponent comp in prj.VBComponents)
    {
        code = comp.CodeModule;

        // Put the name of the code module at the top
        composedFile = comp.Name + Environment.NewLine;

        // Loop through the (1-indexed) lines
        for (int i = 0; i < code.CountOfLines; i++)
        {
            composedFile += code.get_Lines(i + 1, 1) + Environment.NewLine;
        }

        // Add the macro to the list
        macros.Add(composedFile);
    }

    CloseDoc(doc);

    return macros;
}