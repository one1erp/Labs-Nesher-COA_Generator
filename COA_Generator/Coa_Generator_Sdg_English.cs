using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Common;
using DAL;
using Generate_COA_document;
using LSEXT;
using LSSERVICEPROVIDERLib;

namespace COA_Generator
{


    [ComVisible(true)]
    [ProgId("COA_Generator.Coa_Generator_Sdg_English")]
    public class Coa_Generator_Sdg_English : IEntityExtension
    {

        private const bool ISENGLISH = true;
        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            try
            {
                INautilusServiceProvider sp = Parameters["SERVICE_PROVIDER"];

                var records = Parameters["RECORDS"];

                var sdgIds = new List<string>();
                var common = new COAOperation(sp);
                while (!records.EOF)
                {
                    var sdgId = records.Fields["SDG_ID"].Value;
                    sdgIds.Add(sdgId);
                    Sdg sdg = common.GetSdg(long.Parse(sdgId));
                    if (sdg.Status != "C")
                    {
                        return ExecuteExtension.exDisabled;
                    }
                    foreach (var sample in sdg.Samples)
                    {
                        foreach (var aliq in sample.Aliqouts)
                        {
                            if (aliq.Name.Contains("CHL"))
                            {
                                common.Close();
                                return ExecuteExtension.exDisabled;
                            }
                        }
                    }
                    records.MoveNext();
                }
                common.Close();
                return ExecuteExtension.exEnabled;
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                return ExecuteExtension.exDisabled;
            }
        }

        public void Execute(ref LSExtensionParameters Parameters)
        {
  
            try
            {
                INautilusServiceProvider sp = Parameters["SERVICE_PROVIDER"];

                var records = Parameters["RECORDS"];

                var sdgIds = new List<string>();

                while (!records.EOF)
                {
                    var sdgId = records.Fields["SDG_ID"].Value;
                    sdgIds.Add(sdgId);
                    records.MoveNext();
                }

                GenerateCOAforSDg(sdgIds, sp);
            }
            catch (Exception ex)
            {

                Logger.WriteLogFile(ex);

            }
        }

        public void GenerateCOAforSDg(List<string> sdgIds, INautilusServiceProvider sp)
        {
            var common = new COAOperation(sp);
            foreach (var sdgId in sdgIds)
            {
                //Get specified sdg.
                Sdg sdg = common.GetSdg(long.Parse(sdgId));

                //Login new COA report by xml processor 
                var loginCoa = common.LoginCOA("Regular COA 1", sdg, false);
                if (loginCoa)
                {
                    //Retrieve new COA from nautilus.
                    var newCoa = common.GetNewCoaName();
                    //Updates other data
                    common.UpdateNewRegularCoa(newCoa, long.Parse(sdgId), ISENGLISH);
                }
                else
                {
                    One1.Controls.CustomMessageBox.Show("יצירת תעודה נכשלה , אנא פנה לתמיכה.", MessageBoxButtons.OK,
                                                       MessageBoxIcon.Error);
                }
            }
        }
    }
}
