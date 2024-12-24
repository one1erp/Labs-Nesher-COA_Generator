using COA_Generator;
using Common;
using DAL;
using Generate_COA_document;
using Generate_COA_document_V2;
using LSEXT;
using LSSERVICEPROVIDERLib;
using One1.Controls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace COA_Generator
{
    [ComVisible(true), ProgId("COA_Generator.COA_GeneratorAllSamples")]
    public class CoaHelper
    {

        public static void CreateCoaEntity(string sdgId, COAOperation common , List<CoaParameters> coaParametersList, bool IsEng)
        {
            //Get specified sdg.
            Sdg sdg = common.GetSdg(long.Parse(sdgId));

            var loginCoa = common.LoginCOA_V2("Regular COA 1", sdg, false, sdg.LimsGroup.Name);

            if (loginCoa)
            {
                var newCoa = common.GetNewCoaName();

                var coaParams = new CoaParameters
                {
                    NewCoa = newCoa,
                    SdgId = long.Parse(sdgId),
                    IsEnglish = IsEng
                };

                coaParametersList.Add(coaParams);
            }
            else
            {
                One1.Controls.CustomMessageBox.Show("יצירת תעודה נכשלה , אנא פנה לתמיכה.", MessageBoxButtons.OK,
                                                   MessageBoxIcon.Error);
            }
        }
    }
}