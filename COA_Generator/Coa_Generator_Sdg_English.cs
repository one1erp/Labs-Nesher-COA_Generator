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
using Generate_COA_document_V2;

namespace COA_Generator
{


    [ComVisible(true)]
    [ProgId("COA_Generator.Coa_Generator_Sdg_English")]
    public class Coa_Generator_Sdg_English : IEntityExtension
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
            try
            {

                List<CoaParameters> localCoaList = new List<CoaParameters>();
                string sdgIdsStr = string.Join(", ", sdgIds);
                Logger.WriteInfoToLog($"GenerateCOAforSDg ENG called with SDG IDs: {sdgIdsStr}");
                

                var common = new COAOperation(sp);
                TimeHelper.MeasureExecutionTime(() =>
                {
                    foreach (var sdgId in sdgIds)
                    {
                        CoaHelper.CreateCoaEntity(sdgId, common, localCoaList, true);
                    }
                }, "foreach CreateCoaEntity");

                Logger.WriteInfoToLog("coaParametersList : " + coaParametersList.Count.ToString());

                TimeHelper.MeasureExecutionTime(() =>
                {
                    foreach (var coa in localCoaList)
                    {
                        common.UpdateNewRegularCoa(coa.NewCoa, coa.SdgId, true);
                    }
                }, "common.UpdateNewRegularCoa");
                localCoaList.Clear();
            }
            catch (Exception ex) { Logger.WriteExceptionToLog(ex); }
        }
    }
}
