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

namespace COA_Generator
{


    [ComVisible(true)]
    [ProgId("COA_Generator.COA_Generator_Sample")]
    public class COA_Generator_Sample : IEntityExtension
    {
       
        private const bool ISENGLISH = false;
        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {

            try
            {


                var samples = new List<Sample>();
                INautilusServiceProvider sp = Parameters["SERVICE_PROVIDER"];
              
                var records = Parameters["RECORDS"];
                var samplesIDs = new List<string>();
                var common = new COAOperation(sp);
                while (!records.EOF)
                {
                    var sampleID = records.Fields["SAMPLE_ID"].Value;
                    samplesIDs.Add(sampleID);
                    Sample sample = common.GetSample(long.Parse(sampleID.ToString()));
                    samples.Add(sample);
                    if (sample.Status != "C" && sample.Status != "P")
                    {
                        return ExecuteExtension.exDisabled;
                    }

                    foreach (var aliq in sample.Aliqouts)
                    {
                        if (aliq.Name.Contains("CHL"))
                        {
                            common.Close();
                            return ExecuteExtension.exDisabled;
                        }
                    }

                    records.MoveNext();
                }
                return samples.Count > 0 && samples.First().SdgId != null && CheckBase(samples) ? ExecuteExtension.exEnabled : ExecuteExtension.exDisabled;

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
              
                var samplesIDs = new List<string>();

                var common = new COAOperation(sp);
                while (!records.EOF)
                {
                    var sampleID = records.Fields["SAMPLE_ID"].Value;
                    samplesIDs.Add(sampleID);
                    records.MoveNext();
                }
                //Get sdg
                var sdgName = common.GetSample(long.Parse(samplesIDs.First())).Sdg;
                bool b = false;
                if (common.GetSample(long.Parse(samplesIDs.First())).Status == "C")
                {
                    b = common.LoginCOA("COA_GeneratorAllSamples", sdgName, false, common.GetSample(long.Parse(samplesIDs.First())).Name);
                }
                else if (common.GetSample(long.Parse(samplesIDs.First())).Status == "P")
                {
                    b = common.LoginCOA("COA_GeneratorAllSamples", sdgName, true, common.GetSample(long.Parse(samplesIDs.First())).Name);
                    //b = common.LoginCOA("Partial COA", sdg, true, sample.Name);   
                }
                //var b = common.LoginCOA("COA_GeneratorAllSamples", sdgName, true, common.GetSample(long.Parse(samplesIDs.First())).Name);
                //var b = common.LoginCOA("Partial COA", sdgName, true);  
                if (!b)
                {
                    One1.Controls.CustomMessageBox.Show("יצירת תעודה נכשלה , אנא פנה לתמיכה.", MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error);
                }
                var newCOA = common.GetNewCoaName();
                common.UpdateNewPartialCoa(newCOA, samplesIDs, ISENGLISH);
            }
            catch (Exception ex)
            {

                Logger.WriteLogFile(ex);
               
            }

        }






    }
}
