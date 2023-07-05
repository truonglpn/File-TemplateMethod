using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FileWorker.ConcreteTemplate
{
    public class JSON : FileTemplate
    {
        public override DataTable ExtractData(string path)
        {
            return ConvertJsonToDataTable(path);
        }

      
        public DataTable ConvertJsonToDataTable(string jsonData)
        {
            try
            {
                return JsonConvert.DeserializeObject<DataTable>(jsonData);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
    }
}
