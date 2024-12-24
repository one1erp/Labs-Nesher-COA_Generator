using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Common;
using DAL;
using Generate_COA_document;
using Generate_COA_document_V2;
using LSSERVICEPROVIDERLib;
using One1.Controls;
using XmlService;


namespace COA_Generator
{
    public class COAOperation
    {
        #region Fields

        private IDataLayer dal;
        private INautilusServiceProvider sp;
        private NautilusUser user;
        private INautilusProcessXML xmlProcessor;
        private LoginXmlHandler login;

        #endregion

        #region Ctor
        public COAOperation(INautilusServiceProvider sp)
        {

            this.sp = sp;
            xmlProcessor = Utils.GetXmlProcessor(sp);
            var ntlsCon = Utils.GetNtlsCon(sp);
            Utils.CreateConstring(ntlsCon);
            user = Utils.GetNautilusUser(sp);
            dal = new DataLayer();
            dal.Connect();
        }
        #endregion

        #region XML processor

        public bool LoginCOA(string workflowName, Sdg sdg, bool isPartial, string sampleName = "")
        {
            if(workflowName == "Regular COA 1")
            {
                return LoginCOA_V2("Regular COA 1", sdg, false, sdg.LimsGroup.Name);                
            }
            login = new LoginXmlHandler(sp);
            login.CreateLoginXml("U_COA_REPORT", workflowName);
            login.AddProperties("U_SDG", sdg.Name);
            login.AddProperties("U_SAMPLES", sampleName);
            login.AddProperties("U_PARTIAL", isPartial ? "T" : "F");
            login.AddProperties("GROUP_ID", sdg.LimsGroup.Name);
            var s = login.ProcssXml();
            if (!s)
            { Logger.WriteLogFile(login.ErrorResponse, true); }
            return s;
        }

        public bool LoginCOA_V2(string workflowName, Sdg sdg, bool isPartial, string sampleName = "")
        {
            // יצירת אובייקט login אם הוא לא קיים
            if (login == null)
            {
                login = new LoginXmlHandler(sp);
            }

            // אם יש צורך, עדכון פרטי ה-login
            login.CreateLoginXml("U_COA_REPORT", workflowName);
            login.AddProperties("U_SDG", sdg.Name);
            login.AddProperties("U_SAMPLES", sampleName);
            login.AddProperties("U_PARTIAL", isPartial ? "T" : "F");
            login.AddProperties("GROUP_ID", sdg.LimsGroup.Name);

            // קריאה לפונקציה שמעבדת את ה-XML
            bool processResult = login.ProcssXml_V2();
            if (!processResult)
            {
                Logger.WriteLogFile(login.ErrorResponse, true);
            }
            return processResult;
        }


        public COA_Report GetNewCoaName()
        {
            string newName = login.GetValueByTagName("NAME");
            var coa = dal.GetCoaReportByName(newName);
            return coa;
        }

        #endregion

        #region Update new COA
        public Sdg GetSdg(long sdgId)
        {
            return dal.GetSdgById(sdgId);
        }
        
        public Sample GetSample(long sampleId)
        {

            return dal.GetSampleByKey(sampleId);
        }
        private void UpdateCoa(COA_Report coa, bool EN = false)
        {
            coa.Status = "C";
            coa.CreatedOn = DateTime.Now;
            coa.CreatedBy = Convert.ToInt64(user.GetOperatorId());
            coa.U_IS_ENGLISH = EN ? "T" : "F";
        }
        internal void UpdateNewRegularCoa(COA_Report newCOA, long sdgId, bool isENg)
        {
            try
            {
                UpdateCoa(newCOA, isENg);
                newCOA.Partial = "F";
                newCOA.Charge = "F";
                Sdg sdg = GetSdg(sdgId);
                newCOA.ClientId = sdg.SdgClientId;
                newCOA.SdgId = sdg.SdgId;

                
                if (UseNewCoaGeneration())
                {
                    var generateDoc = new Generate_COA_document_V2.GenerateDoc(sdg, dal, newCOA.Name, isENg, false);
                    newCOA.DocPath = generateDoc.OutputPath;
                }
                else
                {
                    var generateDoc = new Generate_COA_document.GenerateDoc(sdg, dal, newCOA.Name, isENg);
                    newCOA.DocPath = generateDoc.SavedPath;
                }
     
                dal.CancelOldCoa(newCOA.Name);
                newCOA.Sdg.COACreated = "T";
                //CustomMessageBox.Show("הופקה תעודה");
            }
            catch (Exception ex)
            {
                CustomMessageBox.Show("מסמך תעודה נכשל", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Logger.WriteLogFile(ex);
                newCOA.Status = "X";
            }

            SaveAndClose();
        }
        internal void UpdateChallenge(COA_Report newCOA, long sdgId, bool Hebrew)
        {
            UpdateCoa(newCOA);
            newCOA.Partial = "F";
            newCOA.Charge = "F";
            Sdg sdg = GetSdg(sdgId);
            newCOA.ClientId = sdg.SdgClientId;
            newCOA.SdgId = sdg.SdgId;
            //Create word document 
            try
            {
                var generateDoc = new GenerateChallengeDoc(sdg, dal, newCOA.Name);
                newCOA.DocPath = generateDoc.SavedPath;
                dal.CancelOldCoa(newCOA.Name);
                newCOA.Sdg.COACreated = "T";
                //CustomMessageBox.Show("הופקה תעודה");
            }
            catch (Exception ex)
            {
                CustomMessageBox.Show("מסמך תעודה נכשל", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Logger.WriteLogFile(ex);
                //Cancel COA if generate doc is failed
                newCOA.Status = "X";


            }
            SaveAndClose();

        }

        public static bool UseNewCoaGeneration()
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            ExeConfigurationFileMap map = new ExeConfigurationFileMap();
            map.ExeConfigFilename = assemblyPath + ".config";
            Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
            var setting = cfg.AppSettings.Settings["UseNewCoaGeneration"];
            //Logger.WriteInfoToLog($"UseNewCoaGeneration : {setting.Value}");
            var use = setting != null && setting.Value == "T";

            return use;
        }
        internal void UpdateNewPartialCoa(COA_Report newCOA, List<string> samplesIDs, bool english)
        {
            List<Sample> samples = new List<Sample>();
            UpdateCoa(newCOA, english);
            newCOA.Charge = "T";
            Sample sample = GetSample(long.Parse(samplesIDs.First()));
            samples.Add(sample);
            if (sample.Sdg != null)
            {
                newCOA.ClientId = newCOA.Sdg.SdgClientId;//Fix bug when sdg client and sample client aren't equal 6/2/24
                newCOA.SdgId = sample.SdgId;
            }
            try
            {
                if (UseNewCoaGeneration())
                {
                    var generateDoc = new Generate_COA_document_V2.GenerateDoc(samples, dal, newCOA.Name, english);
                    newCOA.DocPath = generateDoc.OutputPath;
                }
                else
                {
                    var generateDoc = new Generate_COA_document.GenerateDoc(samples, dal, newCOA.Name, english);
                    newCOA.DocPath = generateDoc.SavedPath;
                }

                dal.CancelOldCoa(newCOA.Name);
                //CustomMessageBox.Show("הופקה תעודה");
            }
            catch (Exception exception)
            {

                CustomMessageBox.Show("מסמך תעודה נכשל", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Logger.WriteLogFile(exception);
                newCOA.Status = "X";

            }

            SaveAndClose();
        }
        internal void UpdateNewSdgBySampleCoa(COA_Report newCOA, long sdgID, bool english)
        {

            UpdateCoa(newCOA, english);

            newCOA.Charge = "T";
            Sdg sdg = GetSdg(sdgID);

            if (sdg != null)
            {
                newCOA.ClientId = newCOA.Sdg.SdgClientId;
                newCOA.SdgId = sdg.SdgId;
            }

            //Generate document
            try
            {

                if (UseNewCoaGeneration())
                {
                    var generateDoc = new Generate_COA_document_V2.GenerateDoc(sdg, dal, newCOA.Name, english, true);
                    newCOA.DocPath = generateDoc.OutputPath;
                }
                else
                {
                    var generateDoc = new GenerateDocPartial(sdg, dal, newCOA.Name, english);
                    newCOA.DocPath = generateDoc.SavedPath;
                }
                
                dal.CancelOldCoa(newCOA.Name);
                newCOA.Sdg.COACreated = "T";
                CustomMessageBox.Show("הופקה תעודה");
                SaveAndClose();
            }
            catch (Exception exception)
            {
                CustomMessageBox.Show("מסמך תעודה נכשל", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Logger.WriteLogFile(exception);
                newCOA.Status = "X";
            }
            
        }
        private void SaveAndClose()
        {
            if (dal != null) dal.SaveChanges();
            Close();
        }

        internal void Close()
        {
            if (dal != null) dal.Close();
        }

        #endregion


    }
}
