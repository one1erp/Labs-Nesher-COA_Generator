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
    [ProgId("COA_Generator.COA_Challenge")]
    public class COA_Challenge : IEntityExtension
    {

        private const bool ISENGLISH = false;
         
        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {

            INautilusServiceProvider sp = Parameters["SERVICE_PROVIDER"];
            var records = Parameters["RECORDS"];
            var common = new COAOperation(sp);
            var sdgId = records.Fields["SDG_ID"].Value;
            Sdg sdg = common.GetSdg(long.Parse(sdgId));
            foreach (var sample in sdg.Samples)
            {
                foreach (var aliq in sample.Aliqouts)
                {
                    if (aliq.Name.Contains("CHL"))
                    {
                        common.Close();
                        return ExecuteExtension.exEnabled;
                    }
                }
            }
            return ExecuteExtension.exDisabled;
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

                GenerateCOAChallenge(sdgIds, sp);
            }
            catch (Exception ex)
            {

                Logger.WriteLogFile(ex);

            }
        }

        public void GenerateCOAChallenge(List<string> sdgIds, INautilusServiceProvider sp)
        {
            var common = new COAOperation(sp);
            foreach (var sdgId in sdgIds)
            {
                //Get specified sdg.
                Sdg sdg = common.GetSdg(long.Parse(sdgId));

                //Login new COA report by xml processor 
                var loginCoa = common.LoginCOA("Regular COA 1", sdg,false);
                if (loginCoa)
                {
                    //Retrieve new COA from nautilus.
                    var newCoa = common.GetNewCoaName();
                    //Updates other data
                    common.UpdateChallenge(newCoa, long.Parse(sdgId), ISENGLISH);
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
