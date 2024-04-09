using System.Windows.Forms;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;


namespace COA_Generator
{

    [ComVisible(true)]
    [ProgId("COA_Generator.COA_Generator")]
    public class COA_Generator : IWorkflowExtension//לא משתמשים עד התשובה מ THERMO
    {
        //Not Active
        INautilusServiceProvider sp;


        public void Execute(ref LSExtensionParameters Parameters)
        {

            //sp = Parameters["SERVICE_PROVIDER"];

            //var records = Parameters["RECORDS"];
            //for (int i = 0; i < records.Fields.Count; i++)
            //{
            //    var x = records.Fields[i].Value;
            //    var x1 = records.Fields[i].Name;
            //}
            //var sdgID = records.Fields["SDG_ID"].Value;

            //var common = new COAOperation(sp);
            //Sdg sdg = common.GetSdg(sdgID.ToString());
            //if (sdg.Status == "A")
            //{
            //    var success = common.LoginCOA("Regular COA",sdg.Name,false);
            //    if (success)
            //    {
            //        var newCOA = common.GetNewCoaName();
            //        common.UpdateNewRegularCoa(newCOA, sdgID.ToString());
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Sdg is  not Authorized");
            //}
        }


    }
}
