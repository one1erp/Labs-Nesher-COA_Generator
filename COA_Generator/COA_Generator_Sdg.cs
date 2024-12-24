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
    [ProgId("COA_Generator.COA_Generator_Sdg")]
    public class COA_Generator_Sdg : IEntityExtension
    {
        List<CoaParameters> coaParametersList = new List<CoaParameters>();
        private COAOperation common;
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

        /// <summary>
        /// </summary>
        /// <param name="sdgIds"></param>
        /// <param name="sp"></param>
        /// 

        //////////////////////////
        public void GenerateCOAforSDg(List<string> sdgIds, INautilusServiceProvider sp)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                common = new COAOperation(sp);

                TimeHelper.MeasureExecutionTime(() =>
                {
                    foreach (var sdgId in sdgIds)
                    {
                        CoaHelper.CreateCoaEntity(sdgId, common, coaParametersList,false);
                    }
                }, "foreach CreateCoaEntity");

                TimeHelper.MeasureExecutionTime(() =>
                {
                    foreach (var coa in coaParametersList)
                    {
                        common.UpdateNewRegularCoa(coa.NewCoa, coa.SdgId, coa.IsEnglish);
                    }
                }, "common.UpdateNewRegularCoa");

            }
            catch (Exception ex)
            {
                Logger.WriteExceptionToLog(ex);
                throw;
            }

            finally
            {
                Cursor.Current = Cursors.Default;
            }

        }
  
    }

    public class CoaParameters
    {
        public COA_Report NewCoa { get; set; }
        public long SdgId { get; set; }
        public bool IsEnglish { get; set; }
    }

    // רשימה לאחסון הפרמטרים
}
