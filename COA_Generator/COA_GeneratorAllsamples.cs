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
using One1.Controls;
using System.Diagnostics;
using Generate_COA_document_V2;

namespace COA_Generator
{


    [ComVisible(true)]
    [ProgId("COA_Generator.COA_GeneratorAllSamples")]
    public class COA_GeneratorAllSamples : CoaHelper, IEntityExtension
    {

        internal static List<CoaParameters> coaParametersList = new List<CoaParameters>();

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
                    if (sdg.Status != "P")
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
            try
            {
                var common = new COAOperation(sp);
                TimeHelper.MeasureExecutionTime(() =>
                {
                    foreach (var sdgId in sdgIds)
                    {
                        CreateCoaEntity(sdgId, common, coaParametersList, false);
                    }
                }, "foreach CreateCoaEntity");

                TimeHelper.MeasureExecutionTime(() =>
                {
                    foreach (var coa in coaParametersList)
                    {
                        common.UpdateNewSdgBySampleCoa(coa.NewCoa, coa.SdgId, coa.IsEnglish);
                    }
                }, "common.UpdateNewSdgBySampleCoa");
                MessageBox.Show($"������ ����� ������");
            }
            catch (Exception ex) {

                Logger.WriteExceptionToLog(ex);
            }
        }
    }
}
