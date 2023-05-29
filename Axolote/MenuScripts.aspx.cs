using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using EspaceJSON;

namespace Axolote
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnProduireJSON_Click(object sender, EventArgs e)
        {
            EspaceJSON.Program.ConvertirExcelEnJson("Fichier");
        }
    }
}