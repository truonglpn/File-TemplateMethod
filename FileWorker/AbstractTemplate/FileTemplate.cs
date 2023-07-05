using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FileWorker
{
    public abstract class FileTemplate
    {
        public string FilePath { get; set; }

        public virtual OpenFileDialog OpenDialogStream(OpenFileDialog theDialog, string titleDialog, string nameFilter, string filterOption, string initialDirectionOpenFile)
        {
            theDialog = new OpenFileDialog();
            theDialog.Title = titleDialog;
            theDialog.Filter = string.Format("{0}|*.{1}", nameFilter, filterOption);
            theDialog.InitialDirectory = initialDirectionOpenFile;
            return theDialog;
        }
        public virtual string StreamPathFromDialog(OpenFileDialog theDialog)
        {
            Stream streamFile = null;
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((streamFile = theDialog.OpenFile()) != null)
                    {
                        using (streamFile) // using construct close when leave by exception is thrown
                        {
                            string filename = theDialog.FileName;
                            return filename;
                        }
                    }
                    else
                        return null;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return null;
                }
            }
            else
                return null;
        }

        public virtual void OpenFile(string path)
        {
            Console.WriteLine($"===== Curent try open file {path} =====");
        }

        public abstract DataTable ExtractData(string path); // extract data to analysis


        public virtual void CloseFile(string path)
        {
            Console.WriteLine($"===== Close file {path} =====");
        }

        public DataTable DataMiner(OpenFileDialog theDialog, string titleDialog, string nameFilter, string filterOption, string initialDirectionOpenFile)
        {
            theDialog = OpenDialogStream(theDialog, titleDialog, nameFilter, filterOption, initialDirectionOpenFile); // open dialog
            string path = StreamPathFromDialog(theDialog); // get path
            OpenFile(path); // open file from path stream === ( emulator  )
            return ExtractData(path);
        }
    }
}
