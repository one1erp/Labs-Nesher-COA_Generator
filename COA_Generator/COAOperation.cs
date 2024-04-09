using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Common;
using DAL;
using Generate_COA_document;
//using KasefetFile;
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
            UpdateCoa(newCOA,isENg);
            newCOA.Partial = "F";
            newCOA.Charge = "F";
            Sdg sdg = GetSdg(sdgId);
            newCOA.ClientId = sdg.SdgClientId;
            newCOA.SdgId = sdg.SdgId;

            //Create word document 
            try
            {



                var generateDoc = new GenerateDoc(sdg, dal, newCOA.Name, isENg);
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


        internal void UpdateNewPartialCoa(COA_Report newCOA, List<string> samplesIDs, bool english)
        {
            List<Sample> samples = new List<Sample>();
            UpdateCoa(newCOA,english);
            //  newCOA.Partial = "T";//26.3.15 הילה ואשי לא ברור למה זה ככה
            newCOA.Charge = "T";
            Sample sample = GetSample(long.Parse(samplesIDs.First()));
            samples.Add(sample);
            if (sample.Sdg != null)
            {
                newCOA.ClientId = newCOA.Sdg.SdgClientId;//Fix bug when sdg client and sample client aren't equal 6/2/24
                newCOA.SdgId = sample.SdgId;
            }
            //הילה , הורדנו את השמת הערך לשדה samples על התעודה 26.3.15
            //foreach (var samplesID in samplesIDs)
            //{
            //    newCOA.SamplesIds += samplesID + ",";
            //}
            //--------------------------------
            //Generate document
            try
            {
                var generateDoc = new GenerateDoc(samples, dal, newCOA.Name, english);
                newCOA.DocPath = generateDoc.SavedPath;
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

            UpdateCoa(newCOA,english);

            newCOA.Charge = "T";
            Sdg sdg = GetSdg(sdgID);

            if (sdg != null)
            {
                //Fix bug when sdg client and sample client aren't equal 6/2/24
                //-----------------------------------------------------
                // var firstOrDefault = sdg.Samples.FirstOrDefault();
                //if (firstOrDefault != null) newCOA.ClientId = firstOrDefault.CLIENT_ID;
                newCOA.ClientId = newCOA.Sdg.SdgClientId;
                //-----------------------------------------------------

                newCOA.SdgId = sdg.SdgId;
            }


            //Generate document
            try
            {
                var generateDoc = new GenerateDocPartial(sdg, dal, newCOA.Name, english);
                newCOA.DocPath = generateDoc.SavedPath;
                dal.CancelOldCoa(newCOA.Name);
                newCOA.Sdg.COACreated = "T";
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
