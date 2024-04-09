using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Diagnostics;

namespace COA_Generator
{


    [ComVisible(true)]
    [ProgId("COA_Generator.COA_forEvrySample")]
    public class COA_forEvrySample : IEntityExtension
    {

        private const bool ISENGLISH = false;
        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {

            try
            {


                var samples = new List<Sample>();
                INautilusServiceProvider sp = Parameters["SERVICE_PROVIDER"];

                var records = Parameters["RECORDS"];
                var sdgIDList = new List<string>();
                var common = new COAOperation(sp);
                while (!records.EOF)
                {
                    var sdgId = records.Fields["SDG_ID"].Value;
                    sdgIDList.Add(sdgId);
                    Sdg sdg = common.GetSdg(long.Parse(sdgId.ToString()));
                    // sdgIDList.Add(sdg);
                    if (sdg.Status != "C")
                    {
                        return ExecuteExtension.exDisabled;
                    }

                    //foreach (var aliq in sample.Aliqouts)
                    //{
                    //    if (aliq.Name.Contains("CHL"))
                    //    {
                    //        common.Close();
                    //        return ExecuteExtension.exDisabled;
                    //    }
                    //}

                    records.MoveNext();
                }
                return sdgIDList.Count > 0 && sdgIDList.First() != null ? ExecuteExtension.exEnabled : ExecuteExtension.exDisabled;

            }
            catch (Exception ex)
            {

                Logger.WriteLogFile(ex);
                return ExecuteExtension.exDisabled;
            }
        }

        private bool CheckBase(List<Sample> samples)
        {

            //בודק שהם מאותה הזמנה
            var sdgID = samples.First().SdgId;
            return samples.All(sample => sample.SdgId == sdgID);
        }

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
                INautilusServiceProvider sp = Parameters["SERVICE_PROVIDER"];

                var records = Parameters["RECORDS"];


            
                var common = new COAOperation(sp);

                while (!records.EOF)//Ashi - multiple coa in one click
                {
                    var sdgId = records.Fields["SDG_ID"].Value;
                    //sdgIDList.Add(sdgId);
                    //Sdg sdg = common.GetSdg(long.Parse(sdgId.ToString()));
                    //var sampleID = records.Fields["SAMPLE_ID"].Value;
                    //samplesIDs.Add(sampleID);


                    //Get sdg
                    Sdg sdg = common.GetSdg(long.Parse(sdgId));
                    foreach (var sample in sdg.Samples.Where(x => x.Status != "X"))
                    {
                        var samplesIDs = new List<string>();
                        samplesIDs.Add(sample.SampleId.ToString());
                        //    Debugger.Launch();
                        var b = common.LoginCOA("COA_GeneratorAllSamples", sdg, false, sample.Name);

                        if (!b)
                        {
                            One1.Controls.CustomMessageBox.Show("יצירת תעודה נכשלה , אנא פנה לתמיכה.", MessageBoxButtons.OK,
                                                            MessageBoxIcon.Error);
                        }
                        var newCOA = common.GetNewCoaName();
                        common.UpdateNewPartialCoa(newCOA, samplesIDs, ISENGLISH);
                    }
                    records.MoveNext();

                }
            }
            catch (Exception ex)
            {

                Logger.WriteLogFile(ex);

            }

        }






    }
}
