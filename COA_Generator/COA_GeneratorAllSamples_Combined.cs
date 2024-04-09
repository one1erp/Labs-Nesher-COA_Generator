using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace COA_Generator
{
    [ComVisible(true)]
    [ProgId("COA_Generator.COA_GeneratorAllSamples_Combined")]
    public class COA_GeneratorAllSamples_Combined : IEntityExtension
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
                    if (sdg.Status != "P")
                    {
                        return ExecuteExtension.exDisabled;
                    }

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

                    //Get sdg
                    Sdg sdg = common.GetSdg(long.Parse(sdgId));
                    foreach (var sample in sdg.Samples.Where(x => x.Status != "X"))
                    {
                        var samplesIDs = new List<string>();
                        samplesIDs.Add(sample.SampleId.ToString());
                        bool b = false;
                        if (sample.Status == "C")
                        {
                            b = common.LoginCOA("COA_GeneratorAllSamples", sdg, false, sample.Name);
                        }
                        else if (sample.Status == "P" || sample.Status == "V")
                        {
                            b = common.LoginCOA("COA_GeneratorAllSamples", sdg, true, sample.Name);
                            //b = common.LoginCOA("Partial COA", sdg, true, sample.Name);   
                        }
                        else
                        {
                            continue;
                        }

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
