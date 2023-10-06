using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using CrystalDecisions.Shared;
using ClassLibrary_BPC;
using CrystalDecisions.CrystalReports.Engine;
using ClassLibrary_BPC.bpc.controller;
using ClassLibrary_BPC.bpc.model;

namespace WebReporting
{
    public partial class GenarateReport : System.Web.UI.Page
    {

        ReportDocument _RD = new ReportDocument();

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {

                string token = Request.QueryString["token"];
                //labID.Text = token;

                doGenarateReport(token);

            }
            catch { }


            //this.viewer1("CrystalReport1.rpt");
        }

        
        private void doGenarateReport(string Token)
        {
            try
            {
                cls_ctSYSReportjob ctReport = new cls_ctSYSReportjob();
                //-- Step 1 Get Report Job
                cls_SYSReportjob objReportjob = ctReport.getDataByToken(Token);

                if (objReportjob != null)
                {

                    //-- Step 2 Get Whose
                    List<cls_SYSReportjobwhose> objList_whose = ctReport.getDataWhose(objReportjob.reportjob_id.ToString());

                    //-- Step 3 Genarate report
                    switch (objReportjob.reportjob_type)
                    {
                        //TA
                        case "TA001":
                            viewerTA("TA001.rpt", objReportjob, objList_whose);
                            break;
                        case "TA002":
                            viewerTA("TA002.rpt", objReportjob, objList_whose);
                            break;
                        case "TA003":
                            viewerTA("TA003.rpt", objReportjob, objList_whose);
                            break;
                        case "TA004":
                            viewerTA("TA004.rpt", objReportjob, objList_whose);
                            break;
                        case "TA005":
                            viewerTA("TA005.rpt", objReportjob, objList_whose);
                            break;
                        case "TA006":
                            viewerTA("TA006.rpt", objReportjob, objList_whose);
                            break;
                        case "TA007":
                            viewerTA("TA007.rpt", objReportjob, objList_whose);
                            break;
                        case "TA008":
                            viewerTA("TA008.rpt", objReportjob, objList_whose);
                            break;
                        case "TA009":
                            viewerTA("TA009.rpt", objReportjob, objList_whose);
                            break;
                        case "TA010":
                            viewerTA("TA010.rpt", objReportjob, objList_whose);
                            break;
                        case "TA011":
                            viewerTA("TA011.rpt", objReportjob, objList_whose);
                            break;
                        case "TA012":
                            viewerTA("TA012.rpt", objReportjob, objList_whose);
                            break;
                        case "TA013":
                            viewerTA("TA013.rpt", objReportjob, objList_whose);
                            break;
                        case "TA014":
                            viewerTA("TA014.rpt", objReportjob, objList_whose);
                            break;
                        case "TA015":
                            viewerTA("TA015.rpt", objReportjob, objList_whose);
                            break;

                        //TA ADD NEW
                        case "TA2":
                            viewerTA("TA2.rpt", objReportjob, objList_whose);
                            break;



                        //PA
                        case "PA001":
                            viewerPA("PA001.rpt", objReportjob, objList_whose);
                            break;

                        case "PA002":
                            viewerPA("PA002.rpt", objReportjob, objList_whose);
                            break;

                        case "PA003":
                            viewerPA("PA003.rpt", objReportjob, objList_whose);
                            break;

                        //TAX
                        case "PA004":
                            viewerTax("PA004.rpt", objReportjob, objList_whose);
                            break;

                        case "PA005":
                            viewerTax("PA005.rpt", objReportjob, objList_whose);
                            break;

                        case "PA006":
                            viewerTax("PA006.rpt", objReportjob, objList_whose);
                            break;

                        case "PA007":
                            viewerTax("PA007.rpt", objReportjob, objList_whose);
                            break;

                        case "PA008":
                            viewerTax("PA008.rpt", objReportjob, objList_whose);
                            break;
                        //END TAX

                        case "PA009":
                            viewerPA("PA009.rpt", objReportjob, objList_whose);
                            break;

                        case "PA010":
                            viewerPA("PA010.rpt", objReportjob, objList_whose);
                            break;

                        case "PA011":
                            viewerPA("PA011.rpt", objReportjob, objList_whose);
                            break;

                        //SSO
                        case "PA012":
                            viewerSSO("PA012.rpt", objReportjob, objList_whose);
                            break;

                        case "PA013":
                            viewerSSO("PA013.rpt", objReportjob, objList_whose);
                            break;

                        case "PA014":
                            viewerSSO("PA014.rpt", objReportjob, objList_whose);
                            break;

                        case "PA015":
                            viewerSSO("PA015.rpt", objReportjob, objList_whose);
                            break;

                        case "PA016":
                            viewerSSO("PA016.rpt", objReportjob, objList_whose);
                            break;
                        //END SSO

                        case "PA017":
                            viewerPA("PA017.rpt", objReportjob, objList_whose);
                            break;

                        case "PA018":
                            viewerPA("PA018.rpt", objReportjob, objList_whose);
                            break;

                        //NEW PR By KIM
                        case "PR1":
                            viewerPA("PR1.rpt", objReportjob, objList_whose);
                            break;

                        case "PR2":
                            viewerTax91("PR2.rpt", objReportjob, objList_whose);
                            break;

                        case "PR3":
                            viewerPA("PR3.rpt", objReportjob, objList_whose);
                            break;

                        case "PR4":
                            viewerPA("PR4.rpt", objReportjob, objList_whose);
                            break;

                        case "PR6":
                            viewerPaySlip("PR6.rpt", objReportjob, objList_whose);
                            break;

                        default:
                            Response.Redirect("404.aspx?message=" + "Report Not Found");
                            break;

                    }

                    //-- Update Status (Prevent duplicate calls)
                    objReportjob.reportjob_status = "F";
                    //ctReport.updateStatus(objReportjob);
                }
                else
                {
                    Response.Redirect("404.aspx?message=" + "Found an error, check the transaction again.");
                }


            }
            catch (Exception ex)
            {

            }
        }


        public void viewerTA(string rptName, cls_SYSReportjob obj, List<cls_SYSReportjobwhose> objListWhose)
        {
            string strError = "viewer1";
            try
            {
                CrystalReportViewer1.RefreshReport();

                DataSet ds = new DataSet();
                string strPath = Server.MapPath(".\\Report\\" + rptName);
                
                _RD.Load(strPath);

                CrystalDecisions.Shared.ParameterFields paramFields = new ParameterFields();
                CrystalDecisions.Shared.ParameterField param1Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param1Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param2Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param2Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param3Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param3Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param4Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param4Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param5Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param5Range = new ParameterDiscreteValue();
               
                strError = "param1Field.ParameterFieldName";

                param1Field.ParameterFieldName = "Language";
                param1Range.Value = obj.reportjob_language;
                param1Field.CurrentValues.Add(param1Range);
                paramFields.Add(param1Field);

                param2Field.ParameterFieldName = "CompanyCode";
                param2Range.Value = obj.company_code;
                param2Field.CurrentValues.Add(param2Range);
                paramFields.Add(param2Field);

                param3Field.ParameterFieldName = "FromDate";
                param3Range.Value = obj.reportjob_fromdate;
                param3Field.CurrentValues.Add(param3Range);
                paramFields.Add(param3Field);

                param4Field.ParameterFieldName = "ToDate";
                param4Range.Value = obj.reportjob_todate;
                param4Field.CurrentValues.Add(param4Range);
                paramFields.Add(param4Field);

                param5Field.ParameterFieldName = "PrintBy";
                param5Range.Value = obj.created_by;
                param5Field.CurrentValues.Add(param5Range);
                paramFields.Add(param5Field);


                cls_ctConnection objConn = new cls_ctConnection();
                objConn.doConnect();

                string strEmpID = "";
                
                foreach(cls_SYSReportjobwhose model in objListWhose){
                    strEmpID += "'" + model.worker_code + "',";
                }

                if (strEmpID.Length > 0)
                    strEmpID = strEmpID.TrimEnd(',');


                StringBuilder objStr = new StringBuilder();
                DataTable dt = new DataTable();

                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_INITIAL");               
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_INITIAL");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_COMPANY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_COMPANY");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_WORKER");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_SHIFT");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_SHIFT");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_TIMECARD");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND (TIMECARD_WORKDATE BETWEEN '" + obj.reportjob_fromdate.ToString("MM/dd/yyyy") + "' AND '" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "')");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_TIMECARD");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_TIMELEAVE");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND (TIMELEAVE_FROMDATE >= '" + obj.reportjob_fromdate.ToString("MM/dd/yyyy") + "' AND TIMELEAVE_TODATE <= '" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "')");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_TIMELEAVE");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_LEAVE");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_LEAVE");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_DEP");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_DEP");
                ds.Tables.Add(dt);


                objStr = new StringBuilder();
                objStr.Append(" SELECT HRM_MT_WORKER.COMPANY_CODE, HRM_MT_WORKER.WORKER_CODE");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_DATE FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_DATE");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL01 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_LEVEL01");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL02 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_LEVEL02");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL03 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_LEVEL03");
                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPDEP");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_TIMEOT");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND (TIMEOT_WORKDATE BETWEEN '" + obj.reportjob_fromdate.ToString("MM/dd/yyyy") + "' AND '" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "')");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_TIMEOT");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_TIMESHIFT");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND (TIMESHIFT_WORKDATE BETWEEN '" + obj.reportjob_fromdate.ToString("MM/dd/yyyy") + "' AND '" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "')");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_TIMESHIFT");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_TIMEDAYTYPE");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND (TIMEDAYTYPE_WORKDATE BETWEEN '" + obj.reportjob_fromdate.ToString("MM/dd/yyyy") + "' AND '" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "')");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_TIMEDAYTYPE");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPSALARY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPSALARY");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_POSITION");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_POSITION");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT HRM_MT_WORKER.COMPANY_CODE, HRM_MT_WORKER.WORKER_CODE");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPPOSITION_DATE FROM HRM_TR_EMPPOSITION WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPPOSITION_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPPOSITION_DATE DESC), '') AS EMPPOSITION_DATE");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPPOSITION_POSITION FROM HRM_TR_EMPPOSITION WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPPOSITION_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPPOSITION_DATE DESC), '') AS EMPPOSITION_POSITION");
                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPPOSITION");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_REASON");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_REASON");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_WAGEDAY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND (WAGEDAY_DATE BETWEEN '" + obj.reportjob_fromdate.ToString("MM/dd/yyyy") + "' AND '" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "')");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_WAGEDAY");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_LOCATION");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_LOCATION");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_TIMEONSITE");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND (TIMEONSITE_WORKDATE BETWEEN '" + obj.reportjob_fromdate.ToString("MM/dd/yyyy") + "' AND '" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "')");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_TIMEONSITE");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_PAYDG");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYDG_PAYDATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_PAYDG");
                ds.Tables.Add(dt);


                strError = "RD.SetDataSource";
                _RD.SetDataSource(ds);

                CrystalReportViewer1.EnableParameterPrompt = false;

                strError = "CrystalReportViewer1.ParameterFieldInfo";
                CrystalReportViewer1.ParameterFieldInfo = paramFields;

                strError = "CrystalReportViewer1.ReportSource";
                CrystalReportViewer1.ReportSource = _RD;

                ds.Dispose();


                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {               
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }

        public void viewerPA(string rptName, cls_SYSReportjob obj, List<cls_SYSReportjobwhose> objListWhose)
        {
            string strError = "viewer1";
            try
            {
                CrystalReportViewer1.RefreshReport();

                DataSet ds = new DataSet();
                string strPath = Server.MapPath(".\\Report\\" + rptName);

                _RD.Load(strPath);

                CrystalDecisions.Shared.ParameterFields paramFields = new ParameterFields();
                CrystalDecisions.Shared.ParameterField param1Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param1Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param2Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param2Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param3Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param3Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param4Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param4Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param5Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param5Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param6Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param6Range = new ParameterDiscreteValue();

                strError = "param1Field.ParameterFieldName";

                param1Field.ParameterFieldName = "Language";
                param1Range.Value = obj.reportjob_language;
                param1Field.CurrentValues.Add(param1Range);
                paramFields.Add(param1Field);

                param2Field.ParameterFieldName = "CompanyCode";
                param2Range.Value = obj.company_code;
                param2Field.CurrentValues.Add(param2Range);
                paramFields.Add(param2Field);

                param3Field.ParameterFieldName = "FromDate";
                param3Range.Value = obj.reportjob_fromdate;
                param3Field.CurrentValues.Add(param3Range);
                paramFields.Add(param3Field);

                param4Field.ParameterFieldName = "ToDate";
                param4Range.Value = obj.reportjob_todate;
                param4Field.CurrentValues.Add(param4Range);
                paramFields.Add(param4Field);

                param5Field.ParameterFieldName = "PrintBy";
                param5Range.Value = obj.created_by;
                param5Field.CurrentValues.Add(param5Range);
                paramFields.Add(param5Field);

                param6Field.ParameterFieldName = "PayDate";
                param6Range.Value = obj.reportjob_paydate;
                param6Field.CurrentValues.Add(param6Range);
                paramFields.Add(param6Field);


                cls_ctConnection objConn = new cls_ctConnection();
                objConn.doConnect();

                string strEmpID = "";

                foreach (cls_SYSReportjobwhose model in objListWhose)
                {
                    strEmpID += "'" + model.worker_code + "',";
                }

                if (strEmpID.Length > 0)
                    strEmpID = strEmpID.TrimEnd(',');


                StringBuilder objStr = new StringBuilder();
                DataTable dt = new DataTable();

                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_INITIAL");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_INITIAL");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_COMPANY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_COMPANY");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_WORKER");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_ITEM");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_ITEM");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_DEP");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_DEP");
                ds.Tables.Add(dt);


                objStr = new StringBuilder();
                objStr.Append(" SELECT HRM_MT_WORKER.COMPANY_CODE, HRM_MT_WORKER.WORKER_CODE");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_DATE FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_DATE");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL01 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_LEVEL01");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL02 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_LEVEL02");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL03 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_LEVEL03");
                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPDEP");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_PAYITEM");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYITEM_DATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_PAYITEM");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_PAYTRAN");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYTRAN_PAYDATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_PAYTRAN");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPBANK");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPBANK");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_BANK");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_BANK");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_PAYBANK");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYBANK_PAYDATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_PAYBANK");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_LEVEL");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_LEVEL");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPPROVIDENT");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPPROVIDENT");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPSALARY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPSALARY");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_PAYPF");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYPF_DATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_PAYPF");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPCARD");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPCARD");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_POSITION");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_POSITION");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT HRM_MT_WORKER.COMPANY_CODE, HRM_MT_WORKER.WORKER_CODE");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPPOSITION_DATE FROM HRM_TR_EMPPOSITION WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPPOSITION_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPPOSITION_DATE DESC), '') AS EMPPOSITION_DATE");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPPOSITION_POSITION FROM HRM_TR_EMPPOSITION WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPPOSITION_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPPOSITION_DATE DESC), '') AS EMPPOSITION_POSITION");
                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPPOSITION");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPADDRESS");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPADDRESS");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_PROVINCE");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_PROVINCE");
                ds.Tables.Add(dt);

                strError = "RD.SetDataSource";
                _RD.SetDataSource(ds);

                CrystalReportViewer1.EnableParameterPrompt = false;

                strError = "CrystalReportViewer1.ParameterFieldInfo";
                CrystalReportViewer1.ParameterFieldInfo = paramFields;

                strError = "CrystalReportViewer1.ReportSource";
                CrystalReportViewer1.ReportSource = _RD;

                ds.Dispose();


                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }

        public void viewerTax(string rptName, cls_SYSReportjob obj, List<cls_SYSReportjobwhose> objListWhose)
        {
            string strError = "viewer1";
            try
            {
                CrystalReportViewer1.RefreshReport();

                DataSet ds = new DataSet();
                string strPath = Server.MapPath(".\\Report\\" + rptName);

                _RD.Load(strPath);

                CrystalDecisions.Shared.ParameterFields paramFields = new ParameterFields();
                CrystalDecisions.Shared.ParameterField param1Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param1Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param2Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param2Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param3Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param3Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param4Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param4Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param5Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param5Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param6Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param6Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param7Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param7Range = new ParameterDiscreteValue();

                strError = "param1Field.ParameterFieldName";

                param1Field.ParameterFieldName = "Language";
                param1Range.Value = obj.reportjob_language;
                param1Field.CurrentValues.Add(param1Range);
                paramFields.Add(param1Field);

                param2Field.ParameterFieldName = "CompanyCode";
                param2Range.Value = obj.company_code;
                param2Field.CurrentValues.Add(param2Range);
                paramFields.Add(param2Field);

                param3Field.ParameterFieldName = "FromDate";
                param3Range.Value = obj.reportjob_fromdate;
                param3Field.CurrentValues.Add(param3Range);
                paramFields.Add(param3Field);

                param4Field.ParameterFieldName = "ToDate";
                param4Range.Value = obj.reportjob_todate;
                param4Field.CurrentValues.Add(param4Range);
                paramFields.Add(param4Field);

                param5Field.ParameterFieldName = "PrintBy";
                param5Range.Value = obj.created_by;
                param5Field.CurrentValues.Add(param5Range);
                paramFields.Add(param5Field);

                param6Field.ParameterFieldName = "PayDate";
                param6Range.Value = obj.reportjob_paydate;
                param6Field.CurrentValues.Add(param6Range);
                paramFields.Add(param6Field);

                if (rptName == "PA006.rpt"|| rptName == "PA008.rpt") 
                {
                    param7Field.ParameterFieldName = "MMM";
                    param7Range.Value = "1";
                    param7Field.CurrentValues.Add(param7Range);
                    paramFields.Add(param7Field);
                } else { }


                cls_ctConnection objConn = new cls_ctConnection();
                objConn.doConnect();

                string strEmpID = "";

                foreach (cls_SYSReportjobwhose model in objListWhose)
                {
                    strEmpID += "'" + model.worker_code + "',";
                }

                if (strEmpID.Length > 0)
                    strEmpID = strEmpID.TrimEnd(',');


                StringBuilder objStr = new StringBuilder();
                DataTable dt = new DataTable();

                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_INITIAL");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_INITIAL");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_COMPANY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_COMPANY");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_WORKER");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_COMADDRESS");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND COMBRANCH_CODE = '00000'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_COMADDRESS");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPADDRESS");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPADDRESS");
                ds.Tables.Add(dt);


                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_PAYTRAN");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                if (rptName == "PA005.rpt" || rptName == "PA006.rpt")
                {
                    objStr.Append(" AND PAYTRAN_PAYDATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                }
                else if (rptName == "PA007.rpt" || rptName == "PA008.rpt")
                {}
                else{
                    objStr.Append(" AND PAYTRAN_PAYDATE <= '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                };
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_PAYTRAN");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_COMCARD");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND CARD_TYPE = 'CID'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_COMCARD");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPCARD");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                objStr.Append(" AND CARD_TYPE = 'NTID'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPCARD");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_PROVINCE");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_PROVINCE");
                ds.Tables.Add(dt);

                strError = "RD.SetDataSource";
                _RD.SetDataSource(ds);

                CrystalReportViewer1.EnableParameterPrompt = false;

                strError = "CrystalReportViewer1.ParameterFieldInfo";
                CrystalReportViewer1.ParameterFieldInfo = paramFields;

                strError = "CrystalReportViewer1.ReportSource";
                CrystalReportViewer1.ReportSource = _RD;

                ds.Dispose();


                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }

        public void viewerSSO(string rptName, cls_SYSReportjob obj, List<cls_SYSReportjobwhose> objListWhose)
        {
            string strError = "viewer1";
            try
            {
                CrystalReportViewer1.RefreshReport();

                DataSet ds = new DataSet();
                string strPath = Server.MapPath(".\\Report\\" + rptName);

                _RD.Load(strPath);

                CrystalDecisions.Shared.ParameterFields paramFields = new ParameterFields();
                CrystalDecisions.Shared.ParameterField param1Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param1Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param2Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param2Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param3Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param3Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param4Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param4Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param5Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param5Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param6Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param6Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param7Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param7Range = new ParameterDiscreteValue();

                strError = "param1Field.ParameterFieldName";

                param1Field.ParameterFieldName = "Language";
                param1Range.Value = obj.reportjob_language;
                param1Field.CurrentValues.Add(param1Range);
                paramFields.Add(param1Field);

                param2Field.ParameterFieldName = "CompanyCode";
                param2Range.Value = obj.company_code;
                param2Field.CurrentValues.Add(param2Range);
                paramFields.Add(param2Field);

                param3Field.ParameterFieldName = "FromDate";
                param3Range.Value = obj.reportjob_fromdate;
                param3Field.CurrentValues.Add(param3Range);
                paramFields.Add(param3Field);

                param4Field.ParameterFieldName = "ToDate";
                param4Range.Value = obj.reportjob_todate;
                param4Field.CurrentValues.Add(param4Range);
                paramFields.Add(param4Field);

                param5Field.ParameterFieldName = "PrintBy";
                param5Range.Value = obj.created_by;
                param5Field.CurrentValues.Add(param5Range);
                paramFields.Add(param5Field);

                param6Field.ParameterFieldName = "PayDate";
                param6Range.Value = obj.reportjob_paydate;
                param6Field.CurrentValues.Add(param6Range);
                paramFields.Add(param6Field);


                cls_ctConnection objConn = new cls_ctConnection();
                objConn.doConnect();

                string strEmpID = "";

                foreach (cls_SYSReportjobwhose model in objListWhose)
                {
                    strEmpID += "'" + model.worker_code + "',";
                }

                if (strEmpID.Length > 0)
                    strEmpID = strEmpID.TrimEnd(',');


                StringBuilder objStr = new StringBuilder();
                DataTable dt = new DataTable();

                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_INITIAL");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_INITIAL");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_COMPANY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_COMPANY");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_WORKER");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_COMADDRESS");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND COMBRANCH_CODE = '00000'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_COMADDRESS");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPADDRESS");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPADDRESS");
                ds.Tables.Add(dt);


                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_PAYTRAN");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYTRAN_PAYDATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_PAYTRAN");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_COMCARD");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_COMCARD");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPCARD");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPCARD");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_MT_PROVINCE");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_PROVINCE");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_COMBRANCH");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_COMBRANCH");
                ds.Tables.Add(dt);

                objStr = new StringBuilder();
                objStr.Append(" SELECT *");
                objStr.Append(" FROM HRM_TR_EMPSALARY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPSALARY");
                ds.Tables.Add(dt);

                strError = "RD.SetDataSource";
                _RD.SetDataSource(ds);

                CrystalReportViewer1.EnableParameterPrompt = false;

                strError = "CrystalReportViewer1.ParameterFieldInfo";
                CrystalReportViewer1.ParameterFieldInfo = paramFields;

                strError = "CrystalReportViewer1.ReportSource";
                CrystalReportViewer1.ReportSource = _RD;

                ds.Dispose();


                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }


        public void viewerPaySlip(string rptName, cls_SYSReportjob obj, List<cls_SYSReportjobwhose> objListWhose)
        {
            string strError = "viewer1";
            try
            {
                CrystalReportViewer1.RefreshReport();

                DataSet ds = new DataSet();
                string strPath = Server.MapPath(".\\Report\\" + rptName);

                _RD.Load(strPath);

                CrystalDecisions.Shared.ParameterFields paramFields = new ParameterFields();
                CrystalDecisions.Shared.ParameterField param1Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param1Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param2Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param2Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param3Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param3Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param4Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param4Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param5Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param5Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param6Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param6Range = new ParameterDiscreteValue();

                strError = "param1Field.ParameterFieldName";

                param1Field.ParameterFieldName = "Language";
                param1Range.Value = obj.reportjob_language;
                param1Field.CurrentValues.Add(param1Range);
                paramFields.Add(param1Field);

                param2Field.ParameterFieldName = "CompanyCode";
                param2Range.Value = obj.company_code;
                param2Field.CurrentValues.Add(param2Range);
                paramFields.Add(param2Field);

                param3Field.ParameterFieldName = "FromDate";
                param3Range.Value = obj.reportjob_fromdate;
                param3Field.CurrentValues.Add(param3Range);
                paramFields.Add(param3Field);

                param4Field.ParameterFieldName = "ToDate";
                param4Range.Value = obj.reportjob_todate;
                param4Field.CurrentValues.Add(param4Range);
                paramFields.Add(param4Field);

                param5Field.ParameterFieldName = "PrintBy";
                param5Range.Value = obj.created_by;
                param5Field.CurrentValues.Add(param5Range);
                paramFields.Add(param5Field);

                param6Field.ParameterFieldName = "PayDate";
                param6Range.Value = obj.reportjob_paydate;
                param6Field.CurrentValues.Add(param6Range);
                paramFields.Add(param6Field);


                cls_ctConnection objConn = new cls_ctConnection();
                objConn.doConnect();

                string strEmpID = "";

                foreach (cls_SYSReportjobwhose model in objListWhose)
                {
                    strEmpID += "'" + model.worker_code + "',";
                }

                if (strEmpID.Length > 0)
                    strEmpID = strEmpID.TrimEnd(',');


                StringBuilder objStr = new StringBuilder();
                DataTable dt = new DataTable();
                //tbMTInitial
                objStr.Append(" SELECT ");
                objStr.Append(" INITIAL_CODE AS InitialID");
                objStr.Append(", INITIAL_NAME_EN AS Name");
                objStr.Append(", INITIAL_NAME_TH AS NameT");

                objStr.Append(" FROM HRM_MT_INITIAL");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTInitial");
                ds.Tables.Add(dt);

                //tbMTCompany
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", COMPANY_NAME_EN AS CompNameE");
                objStr.Append(", COMPANY_NAME_TH AS CompName");

                objStr.Append(" FROM HRM_MT_COMPANY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTCompany");
                ds.Tables.Add(dt);

                //tbMTEmpMain
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" WORKER_CODE AS EmpID");
                objStr.Append(", COMPANY_CODE AS CompID");
                objStr.Append(", WORKER_INITIAL AS InitialID");
                objStr.Append(", WORKER_FNAME_EN AS EmpFName");
                objStr.Append(", WORKER_LNAME_EN AS EmpLName");
                objStr.Append(", WORKER_FNAME_TH AS EmpFNameT");
                objStr.Append(", WORKER_LNAME_TH AS EmpLNameT");
                objStr.Append(", WORKER_EMPTYPE AS EmpType");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL01 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS Level01");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL02 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS Level02");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL03 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS Level03");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPPOSITION_POSITION FROM HRM_TR_EMPPOSITION WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPPOSITION_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPPOSITION_DATE DESC), '') AS PositionID");

                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTEmpMain");
                ds.Tables.Add(dt);

                //tbMTEmpType
                objStr = new StringBuilder();
                objStr.Append(" SELECT DISTINCT ");
                objStr.Append(" WORKER_EMPTYPE AS EmpTypeID");
                objStr.Append(", IIF(WORKER_EMPTYPE = 'M','Monthly','Daily') AS Name");
                objStr.Append(", IIF(WORKER_EMPTYPE = 'M','รายเดือน','รายวัน') AS NameT");
                objStr.Append(" FROM HRM_MT_WORKER");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTEmpType");
                ds.Tables.Add(dt);

                //tbMTAllwDeducType
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", ITEM_CODE AS AllwDeducID");
                objStr.Append(", ITEM_NAME_TH AS AllwDeducDesT");
                objStr.Append(", ITEM_NAME_EN AS AllwDeducDesE");
                
                objStr.Append(" FROM HRM_MT_ITEM");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTAllwDeducType");
                ds.Tables.Add(dt);

                //tbMTPart
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", DEP_LEVEL AS LevelID");
                objStr.Append(", DEP_CODE AS PartID");
                objStr.Append(", DEP_NAME_TH AS PartNameT");
                objStr.Append(", DEP_NAME_EN AS PartNameE");

                objStr.Append(" FROM HRM_MT_DEP");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTPart");
                ds.Tables.Add(dt);


                //objStr = new StringBuilder();
                //objStr.Append(" SELECT HRM_MT_WORKER.COMPANY_CODE, HRM_MT_WORKER.WORKER_CODE");
                //objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_DATE FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_DATE");
                //objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL01 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_LEVEL01");
                //objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL02 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_LEVEL02");
                //objStr.Append(", ISNULL((SELECT TOP 1 EMPDEP_LEVEL03 FROM HRM_TR_EMPDEP WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPDEP_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPDEP_DATE DESC), '') AS EMPDEP_LEVEL03");
                //objStr.Append(" FROM HRM_MT_WORKER");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                //dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPDEP");
                //ds.Tables.Add(dt);

                //vwAllowDeduct
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", WORKER_CODE AS EmpID");
                objStr.Append(", ITEM_CODE AS AllwDeducID");
                objStr.Append(", PAYITEM_DATE AS FromDate");
                objStr.Append(", PAYITEM_DATE AS ToDate");
                objStr.Append(", PAYITEM_AMOUNT AS Amount");
                objStr.Append(", (SELECT ITEM_TYPE FROM HRM_MT_ITEM WHERE COMPANY_CODE = HRM_TR_PAYITEM.COMPANY_CODE AND ITEM_CODE = HRM_TR_PAYITEM.ITEM_CODE) AS TypeAllwOrDe");
                objStr.Append(", PAYITEM_QUANTITY AS QuantityAD");
                objStr.Append(" FROM HRM_TR_PAYITEM");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYITEM_DATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "vwAllowDeduct");
                ds.Tables.Add(dt);

                //tbTRPRPaytran
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", WORKER_CODE AS EmpID");
                objStr.Append(", PAYTRAN_PAYDATE AS PayDate");
                objStr.Append(", PAYTRAN_SSOEMP AS SSO");
                objStr.Append(", (ISNULL(PAYTRAN_TAX_401,0)+ISNULL(PAYTRAN_TAX_4012,0)+ ISNULL(PAYTRAN_TAX_4013,0)+ISNULL(PAYTRAN_TAX_402I,0)+ISNULL(PAYTRAN_TAX_402O,0)) AS Tax");
                objStr.Append(", PAYTRAN_PFEMP AS PFEmpMoney");
                objStr.Append(", PAYTRAN_INCOME_TOTAL AS TotalIncome");
                objStr.Append(", PAYTRAN_DEDUCT_TOTAL AS TotalDeduct");
                objStr.Append(", (ISNULL(PAYTRAN_NETPAY_B,0)+ISNULL(PAYTRAN_NEYPAY_C,0)) AS NetPay");
                objStr.Append(" FROM HRM_TR_PAYTRAN");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYTRAN_PAYDATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "tbTRPRPaytran");
                ds.Tables.Add(dt);

                //tbTREmpBank
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", WORKER_CODE AS EmpID");
                objStr.Append(", EMPBANK_BANKACCOUNT AS BankAccNo");
                objStr.Append(", EMPBANK_BANKNAME AS BankAccName");
                objStr.Append(" FROM HRM_TR_EMPBANK");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "tbTREmpBank");
                ds.Tables.Add(dt);

                //objStr = new StringBuilder();
                //objStr.Append(" SELECT *");
                //objStr.Append(" FROM HRM_MT_BANK");
                //dt = objConn.doGetTable(objStr.ToString(), "HRM_MT_BANK");
                //ds.Tables.Add(dt);

                //objStr = new StringBuilder();
                //objStr.Append(" SELECT *");
                //objStr.Append(" FROM HRM_TR_PAYBANK");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //objStr.Append(" AND PAYBANK_PAYDATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                //objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                //dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_PAYBANK");
                //ds.Tables.Add(dt);

                //tbMTLevel
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", LEVEL_CODE AS LevelID");
                objStr.Append(", LEVEL_NAME_EN AS Name");
                objStr.Append(", LEVEL_NAME_TH AS NameT");
                objStr.Append(" FROM HRM_MT_LEVEL");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTLevel");
                ds.Tables.Add(dt);

                //tbMTPosition
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", POSITION_CODE AS PositionID");
                objStr.Append(", POSITION_NAME_TH AS PositionNameT");
                objStr.Append(", POSITION_NAME_EN AS PositionNameE");
                objStr.Append(" FROM HRM_MT_POSITION");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTPosition");
                ds.Tables.Add(dt);

                //objStr = new StringBuilder();
                //objStr.Append(" SELECT HRM_MT_WORKER.COMPANY_CODE, HRM_MT_WORKER.WORKER_CODE");
                //objStr.Append(", ISNULL((SELECT TOP 1 EMPPOSITION_DATE FROM HRM_TR_EMPPOSITION WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPPOSITION_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPPOSITION_DATE DESC), '') AS EMPPOSITION_DATE");
                //objStr.Append(", ISNULL((SELECT TOP 1 EMPPOSITION_POSITION FROM HRM_TR_EMPPOSITION WHERE COMPANY_CODE=HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE=HRM_MT_WORKER.WORKER_CODE AND EMPPOSITION_DATE<='" + obj.reportjob_todate.ToString("MM/dd/yyyy") + "' ORDER BY EMPPOSITION_DATE DESC), '') AS EMPPOSITION_POSITION");
                //objStr.Append(" FROM HRM_MT_WORKER");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                //dt = objConn.doGetTable(objStr.ToString(), "HRM_TR_EMPPOSITION");
                //ds.Tables.Add(dt);

                //tbTRPRAccumulate
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" a.COMPANY_CODE AS CompID ");
                objStr.Append(", a.WORKER_CODE AS EmpID ");
                objStr.Append(", a.PAYTRAN_PAYDATE AS PayDate ");
                objStr.Append(", b.incomeFix ");
                objStr.Append(", b.Tax ");
                objStr.Append(", b.SocialFix ");
                objStr.Append(", b.SSOAcc_Com ");
                objStr.Append(", b.PfAcc ");
                objStr.Append(", b.PfAcc_Com ");
                objStr.Append(" FROM HRM_TR_PAYTRAN a ");
                objStr.Append(" CROSS JOIN( ");
                objStr.Append(" SELECT ");
                objStr.Append(" SUM((ISNULL(PAYTRAN_INCOME_401,0) - ISNULL(PAYTRAN_DEDUCT_401,0)) + (ISNULL(PAYTRAN_INCOME_4012,0) - ISNULL(PAYTRAN_DEDUCT_4012,0)) + (ISNULL(PAYTRAN_INCOME_4013,0) - ISNULL(PAYTRAN_DEDUCT_4013,0)) + (ISNULL(PAYTRAN_INCOME_402I,0) - ISNULL(PAYTRAN_DEDUCT_402I,0)) + (ISNULL(PAYTRAN_INCOME_402O,0) - ISNULL(PAYTRAN_DEDUCT_402O,0))) AS incomeFix ");
                objStr.Append(", SUM((ISNULL(PAYTRAN_TAX_401,0)+ISNULL(PAYTRAN_TAX_4012,0)+ ISNULL(PAYTRAN_TAX_4013,0)+ISNULL(PAYTRAN_TAX_402I,0)+ISNULL(PAYTRAN_TAX_402O,0))) AS Tax ");
                objStr.Append(", SUM(ISNULL(PAYTRAN_SSOEMP,0)) AS SocialFix ");
                objStr.Append(", SUM(ISNULL(PAYTRAN_SSOCOM,0)) AS SSOAcc_Com ");
                objStr.Append(", SUM(ISNULL(PAYTRAN_PFEMP,0))AS PfAcc ");
                objStr.Append(", SUM(ISNULL(PAYTRAN_PFCOM,0))AS PfAcc_Com");
                objStr.Append(" FROM HRM_TR_PAYTRAN ");
                objStr.Append(" WHERE PAYTRAN_PAYDATE IN (SELECT PERIOD_PAYMENT FROM HRM_MT_PERIOD WHERE YEAR_CODE='" + obj.reportjob_paydate.ToString("yyyy") + "' AND PERIOD_PAYMENT <= '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "' )");
                objStr.Append(" AND COMPANY_CODE = '" + obj.company_code + "' ");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                objStr.Append(" )b ");

                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYTRAN_PAYDATE IN (SELECT PERIOD_PAYMENT FROM HRM_MT_PERIOD WHERE YEAR_CODE='" + obj.reportjob_paydate.ToString("yyyy") + "' AND PERIOD_PAYMENT <= '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "' )");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "tbTRPRAccumulate");
                ds.Tables.Add(dt);

                //tbTRLeaveAccumurate
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", WORKER_CODE AS EmpID");
                objStr.Append(", YEAR_CODE AS LeaveYear");
                objStr.Append(", LEAVE_CODE AS LeaveID");
                objStr.Append(", EMPLEAVEACC_BF AS LeaveBF");
                objStr.Append(", EMPLEAVEACC_ANNUAL AS AnnualLeave");
                objStr.Append(", EMPLEAVEACC_USED AS LeaveUsed");

                objStr.Append(" FROM HRM_TR_EMPLEAVEACC");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                objStr.Append(" AND YEAR_CODE = '" + obj.reportjob_paydate.ToString("yyyy") + "'");

                dt = objConn.doGetTable(objStr.ToString(), "tbTRLeaveAccumurate");
                ds.Tables.Add(dt);


                strError = "RD.SetDataSource";
                _RD.SetDataSource(ds);

                CrystalReportViewer1.EnableParameterPrompt = false;

                strError = "CrystalReportViewer1.ParameterFieldInfo";
                CrystalReportViewer1.ParameterFieldInfo = paramFields;

                strError = "CrystalReportViewer1.ReportSource";
                CrystalReportViewer1.ReportSource = _RD;

                ds.Dispose();


                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }

        public void viewerTax91(string rptName, cls_SYSReportjob obj, List<cls_SYSReportjobwhose> objListWhose)
        {
            string strError = "viewer1";
            try
            {
                CrystalReportViewer1.RefreshReport();

                DataSet ds = new DataSet();
                string strPath = Server.MapPath(".\\Report\\" + rptName);

                _RD.Load(strPath);

                CrystalDecisions.Shared.ParameterFields paramFields = new ParameterFields();
                CrystalDecisions.Shared.ParameterField param1Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param1Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param2Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param2Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param3Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param3Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param4Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param4Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param5Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param5Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param6Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param6Range = new ParameterDiscreteValue();
                CrystalDecisions.Shared.ParameterField param7Field = new ParameterField();
                CrystalDecisions.Shared.ParameterDiscreteValue param7Range = new ParameterDiscreteValue();

                strError = "param1Field.ParameterFieldName";

                param1Field.ParameterFieldName = "Language";
                param1Range.Value = obj.reportjob_language;
                param1Field.CurrentValues.Add(param1Range);
                paramFields.Add(param1Field);

                param2Field.ParameterFieldName = "CompanyCode";
                param2Range.Value = obj.company_code;
                param2Field.CurrentValues.Add(param2Range);
                paramFields.Add(param2Field);

                param3Field.ParameterFieldName = "FromDate";
                param3Range.Value = obj.reportjob_fromdate;
                param3Field.CurrentValues.Add(param3Range);
                paramFields.Add(param3Field);

                param4Field.ParameterFieldName = "ToDate";
                param4Range.Value = obj.reportjob_todate;
                param4Field.CurrentValues.Add(param4Range);
                paramFields.Add(param4Field);

                param5Field.ParameterFieldName = "PrintBy";
                param5Range.Value = obj.created_by;
                param5Field.CurrentValues.Add(param5Range);
                paramFields.Add(param5Field);

                param6Field.ParameterFieldName = "PayDate";
                param6Range.Value = obj.reportjob_paydate;
                param6Field.CurrentValues.Add(param6Range);
                paramFields.Add(param6Field);

                param7Field.ParameterFieldName = "Month";
                param7Range.Value = "1";
                param7Field.CurrentValues.Add(param7Range);
                paramFields.Add(param7Field);


                cls_ctConnection objConn = new cls_ctConnection();
                objConn.doConnect();

                string strEmpID = "";

                foreach (cls_SYSReportjobwhose model in objListWhose)
                {
                    strEmpID += "'" + model.worker_code + "',";
                }

                if (strEmpID.Length > 0)
                    strEmpID = strEmpID.TrimEnd(',');


                StringBuilder objStr = new StringBuilder();
                DataTable dt = new DataTable();
                //tbMTInitial
                objStr.Append(" SELECT ");
                objStr.Append(" INITIAL_CODE AS InitialID");
                objStr.Append(", INITIAL_NAME_EN AS Name");
                objStr.Append(", INITIAL_NAME_TH AS NameT");

                objStr.Append(" FROM HRM_MT_INITIAL");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTInitial");
                ds.Tables.Add(dt);

                //tbMTCompany
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", COMPANY_NAME_EN AS CompNameE");
                objStr.Append(", COMPANY_NAME_TH AS CompName");

                objStr.Append(", (SELECT TOP 1 COMADDRESS_TEL FROM HRM_TR_COMADDRESS WHERE COMPANY_CODE = HRM_MT_COMPANY.COMPANY_CODE) AS CompTelephone");
                objStr.Append(", (SELECT TOP 1 COMCARD_CODE FROM HRM_TR_COMCARD WHERE COMPANY_CODE = HRM_MT_COMPANY.COMPANY_CODE AND CARD_TYPE = 'CID') AS CitizenID");
                objStr.Append(", (SELECT TOP 1 PROVINCE_CODE FROM HRM_TR_COMADDRESS WHERE COMPANY_CODE = HRM_MT_COMPANY.COMPANY_CODE) AS ProvinceID");

                objStr.Append(" FROM HRM_MT_COMPANY");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTCompany");
                ds.Tables.Add(dt);

                //tbMTEmpMain
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" WORKER_CODE AS EmpID");
                objStr.Append(", COMPANY_CODE AS CompID");
                objStr.Append(", (SELECT TOP 1 COMBRANCH_CODE FROM HRM_TR_COMBRANCH WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE) AS BranchID");

                objStr.Append(", WORKER_INITIAL AS InitialID");
                objStr.Append(", WORKER_FNAME_EN AS EmpFName");
                objStr.Append(", WORKER_LNAME_EN AS EmpLName");
                objStr.Append(", WORKER_FNAME_TH AS EmpFNameT");
                objStr.Append(", WORKER_LNAME_TH AS EmpLNameT");
                objStr.Append(", WORKER_EMPTYPE AS EmpType");
                objStr.Append(", WORKER_BIRTHDATE AS BirthDay");

                objStr.Append(", ISNULL((SELECT EMPCARD_CODE FROM HRM_TR_EMPCARD WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE AND CARD_TYPE = 'NTID'),'') AS CardNo");
                objStr.Append(", ISNULL((SELECT EMPCARD_CODE FROM HRM_TR_EMPCARD WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE AND CARD_TYPE = 'TAX'),ISNULL((SELECT EMPCARD_CODE FROM HRM_TR_EMPCARD WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE AND CARD_TYPE = 'NTID'),'')) AS TaxNo");
                objStr.Append(", ISNULL((SELECT EMPFAMILY_CODE FROM HRM_TR_EMPFAMILY WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE AND FAMILY_TYPE = '04' ),'') AS EmpFaCardNo");
                objStr.Append(", ISNULL((SELECT EMPFAMILY_CODE FROM HRM_TR_EMPFAMILY WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE AND FAMILY_TYPE = '09' ),'') AS EmpMoCardNo");

                objStr.Append(", '' AS SpoFaCardNo");
                objStr.Append(", '' AS SpoMoCardNo");
                objStr.Append(", ISNULL((SELECT EMPFAMILY_FNAME_TH FROM HRM_TR_EMPFAMILY WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE AND (FAMILY_TYPE = '02' OR FAMILY_TYPE = '11') ),'') AS SpoFName");
                objStr.Append(", ISNULL((SELECT EMPFAMILY_LNAME_TH FROM HRM_TR_EMPFAMILY WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE AND (FAMILY_TYPE = '02'OR FAMILY_TYPE = '11') ),'') AS SpoLName");
                objStr.Append(", ISNULL((SELECT EMPFAMILY_CODE FROM HRM_TR_EMPFAMILY WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE AND (FAMILY_TYPE = '02'OR FAMILY_TYPE = '11') ),'') AS SpoCardNo");
                objStr.Append(", ISNULL((SELECT EMPFAMILY_BIRTHDATE FROM HRM_TR_EMPFAMILY WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE AND (FAMILY_TYPE = '02'OR FAMILY_TYPE = '11') ),'') AS SpoBirthDay");

                objStr.Append(", '' AS MaritalTax");
                objStr.Append(", '' AS SentTaxStatus");
                objStr.Append(", 0 AS TaxNoStudyChild");
                objStr.Append(", 0 AS TaxStudyChild");

                objStr.Append(", ISNULL((SELECT TOP 1 EMPADDRESS_NO FROM HRM_TR_EMPADDRESS WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE),'' ) AS PreAddrNo");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPADDRESS_MOO FROM HRM_TR_EMPADDRESS WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE),'' ) AS PreMoo");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPADDRESS_SOI FROM HRM_TR_EMPADDRESS WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE),'' ) AS PreSoi");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPADDRESS_ROAD FROM HRM_TR_EMPADDRESS WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE),'' ) AS PreRoad");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPADDRESS_TAMBON FROM HRM_TR_EMPADDRESS WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE),'' ) AS PreTambon");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPADDRESS_AMPHUR FROM HRM_TR_EMPADDRESS WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE),'' ) AS PreAmphur");
                objStr.Append(", ISNULL((SELECT TOP 1 PROVINCE_CODE FROM HRM_TR_EMPADDRESS WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE),'' ) AS PreProvinceID");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPADDRESS_ZIPCODE FROM HRM_TR_EMPADDRESS WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE),'' ) AS PreZipcode");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPADDRESS_TEL FROM HRM_TR_EMPADDRESS WHERE COMPANY_CODE = HRM_MT_WORKER.COMPANY_CODE AND WORKER_CODE = HRM_MT_WORKER.WORKER_CODE),'' ) AS PreTel");

                objStr.Append(" FROM HRM_MT_WORKER");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTEmpMain");
                ds.Tables.Add(dt);

                //tbMTEmpType
                objStr = new StringBuilder();
                objStr.Append(" SELECT DISTINCT ");
                objStr.Append(" WORKER_EMPTYPE AS EmpTypeID");
                objStr.Append(", IIF(WORKER_EMPTYPE = 'M','Monthly','Daily') AS Name");
                objStr.Append(", IIF(WORKER_EMPTYPE = 'M','รายเดือน','รายวัน') AS NameT");
                objStr.Append(" FROM HRM_MT_WORKER");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTEmpType");
                ds.Tables.Add(dt);

                //tbMTAllwDeducType
                //objStr = new StringBuilder();
                //objStr.Append(" SELECT ");
                //objStr.Append(" COMPANY_CODE AS CompID");
                //objStr.Append(", ITEM_CODE AS AllwDeducID");
                //objStr.Append(", ITEM_NAME_TH AS AllwDeducDesT");
                //objStr.Append(", ITEM_NAME_EN AS AllwDeducDesE");

                //objStr.Append(" FROM HRM_MT_ITEM");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //dt = objConn.doGetTable(objStr.ToString(), "tbMTAllwDeducType");
                //ds.Tables.Add(dt);

                //tbMTPart
                //objStr = new StringBuilder();
                //objStr.Append(" SELECT ");
                //objStr.Append(" COMPANY_CODE AS CompID");
                //objStr.Append(", DEP_LEVEL AS LevelID");
                //objStr.Append(", DEP_CODE AS PartID");
                //objStr.Append(", DEP_NAME_TH AS PartNameT");
                //objStr.Append(", DEP_NAME_EN AS PartNameE");

                //objStr.Append(" FROM HRM_MT_DEP");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //dt = objConn.doGetTable(objStr.ToString(), "tbMTPart");
                //ds.Tables.Add(dt);

                //vwAllowDeduct
                //objStr = new StringBuilder();
                //objStr.Append(" SELECT ");
                //objStr.Append(" COMPANY_CODE AS CompID");
                //objStr.Append(", WORKER_CODE AS EmpID");
                //objStr.Append(", ITEM_CODE AS AllwDeducID");
                //objStr.Append(", PAYITEM_DATE AS FromDate");
                //objStr.Append(", PAYITEM_DATE AS ToDate");
                //objStr.Append(", PAYITEM_AMOUNT AS Amount");
                //objStr.Append(", (SELECT ITEM_TYPE FROM HRM_MT_ITEM WHERE COMPANY_CODE = HRM_TR_PAYITEM.COMPANY_CODE AND ITEM_CODE = HRM_TR_PAYITEM.ITEM_CODE) AS TypeAllwOrDe");
                //objStr.Append(", PAYITEM_QUANTITY AS QuantityAD");
                //objStr.Append(" FROM HRM_TR_PAYITEM");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //objStr.Append(" AND PAYITEM_DATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                //objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                //dt = objConn.doGetTable(objStr.ToString(), "vwAllowDeduct");
                //ds.Tables.Add(dt);

                //tbTRPRPaytran
                //objStr = new StringBuilder();
                //objStr.Append(" SELECT ");
                //objStr.Append(" COMPANY_CODE AS CompID");
                //objStr.Append(", WORKER_CODE AS EmpID");
                //objStr.Append(", PAYTRAN_PAYDATE AS PayDate");
                //objStr.Append(", PAYTRAN_SSOEMP AS SSO");
                //objStr.Append(", (ISNULL(PAYTRAN_TAX_401,0)+ISNULL(PAYTRAN_TAX_4012,0)+ ISNULL(PAYTRAN_TAX_4013,0)+ISNULL(PAYTRAN_TAX_402I,0)+ISNULL(PAYTRAN_TAX_402O,0)) AS Tax");
                //objStr.Append(", PAYTRAN_PFEMP AS PFEmpMoney");
                //objStr.Append(", PAYTRAN_INCOME_TOTAL AS TotalIncome");
                //objStr.Append(", PAYTRAN_DEDUCT_TOTAL AS TotalDeduct");
                //objStr.Append(", (ISNULL(PAYTRAN_NETPAY_B,0)+ISNULL(PAYTRAN_NEYPAY_C,0)) AS NetPay");
                //objStr.Append(" FROM HRM_TR_PAYTRAN");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //objStr.Append(" AND PAYTRAN_PAYDATE = '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "'");
                //objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                //dt = objConn.doGetTable(objStr.ToString(), "tbTRPRPaytran");
                //ds.Tables.Add(dt);

                //tbTREmpBank
                //objStr = new StringBuilder();
                //objStr.Append(" SELECT ");
                //objStr.Append(" COMPANY_CODE AS CompID");
                //objStr.Append(", WORKER_CODE AS EmpID");
                //objStr.Append(", EMPBANK_BANKACCOUNT AS BankAccNo");
                //objStr.Append(", EMPBANK_BANKNAME AS BankAccName");
                //objStr.Append(" FROM HRM_TR_EMPBANK");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                //dt = objConn.doGetTable(objStr.ToString(), "tbTREmpBank");
                //ds.Tables.Add(dt);

                //tbMTLevel
                //objStr = new StringBuilder();
                //objStr.Append(" SELECT ");
                //objStr.Append(" COMPANY_CODE AS CompID");
                //objStr.Append(", LEVEL_CODE AS LevelID");
                //objStr.Append(", LEVEL_NAME_EN AS Name");
                //objStr.Append(", LEVEL_NAME_TH AS NameT");
                //objStr.Append(" FROM HRM_MT_LEVEL");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //dt = objConn.doGetTable(objStr.ToString(), "tbMTLevel");
                //ds.Tables.Add(dt);

                //tbMTPosition
                //objStr = new StringBuilder();
                //objStr.Append(" SELECT ");
                //objStr.Append(" COMPANY_CODE AS CompID");
                //objStr.Append(", POSITION_CODE AS PositionID");
                //objStr.Append(", POSITION_NAME_TH AS PositionNameT");
                //objStr.Append(", POSITION_NAME_EN AS PositionNameE");
                //objStr.Append(" FROM HRM_MT_POSITION");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //dt = objConn.doGetTable(objStr.ToString(), "tbMTPosition");
                //ds.Tables.Add(dt);

                //tbTRPRAccumulate
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" a.COMPANY_CODE AS CompID ");
                objStr.Append(", a.WORKER_CODE AS EmpID ");
                objStr.Append(", a.PAYTRAN_PAYDATE AS PayDate ");
                objStr.Append(", b.incomeFix ");
                objStr.Append(", b.Tax ");
                objStr.Append(", b.SocialFix ");
                objStr.Append(", b.SSOAcc_Com ");
                objStr.Append(", b.PfAcc ");
                objStr.Append(", b.PfAcc_Com ");

                objStr.Append(", b.TaxFortyOne");
                objStr.Append(", b.TaxFortyOnePerThree");
                objStr.Append(", b.TaxFortyOneTwo");
                objStr.Append(", b.TaxFortyTwoIn");
                objStr.Append(", b.TaxFortyTwoOut");
                objStr.Append(", b.IncFortyOne");
                objStr.Append(", b.IncFortyOnePerThree");
                objStr.Append(", b.IncFortyOneTwo");
                objStr.Append(", b.IncFortyTwoIn");
                objStr.Append(", b.IncFortyTwoOut");

                objStr.Append(" FROM HRM_TR_PAYTRAN a ");
                objStr.Append(" CROSS JOIN( ");
                objStr.Append(" SELECT ");
                objStr.Append(" SUM((ISNULL(PAYTRAN_INCOME_401,0) - ISNULL(PAYTRAN_DEDUCT_401,0)) + (ISNULL(PAYTRAN_INCOME_4012,0) - ISNULL(PAYTRAN_DEDUCT_4012,0)) + (ISNULL(PAYTRAN_INCOME_4013,0) - ISNULL(PAYTRAN_DEDUCT_4013,0)) + (ISNULL(PAYTRAN_INCOME_402I,0) - ISNULL(PAYTRAN_DEDUCT_402I,0)) + (ISNULL(PAYTRAN_INCOME_402O,0) - ISNULL(PAYTRAN_DEDUCT_402O,0))) AS incomeFix ");
                objStr.Append(", SUM((ISNULL(PAYTRAN_TAX_401,0)+ISNULL(PAYTRAN_TAX_4012,0)+ ISNULL(PAYTRAN_TAX_4013,0)+ISNULL(PAYTRAN_TAX_402I,0)+ISNULL(PAYTRAN_TAX_402O,0))) AS Tax ");
                objStr.Append(", SUM(ISNULL(PAYTRAN_SSOEMP,0)) AS SocialFix ");
                objStr.Append(", SUM(ISNULL(PAYTRAN_SSOCOM,0)) AS SSOAcc_Com ");
                objStr.Append(", SUM(ISNULL(PAYTRAN_PFEMP,0))AS PfAcc ");
                objStr.Append(", SUM(ISNULL(PAYTRAN_PFCOM,0))AS PfAcc_Com");

                objStr.Append(", SUM(ISNULL(PAYTRAN_TAX_401,0)) AS TaxFortyOne");
                objStr.Append(", SUM(ISNULL(PAYTRAN_TAX_4013,0)) AS TaxFortyOnePerThree");
                objStr.Append(", SUM(ISNULL(PAYTRAN_TAX_4012,0)) AS TaxFortyOneTwo");
                objStr.Append(", SUM(ISNULL(PAYTRAN_TAX_402I,0)) AS TaxFortyTwoIn");
                objStr.Append(", SUM(ISNULL(PAYTRAN_TAX_402O,0)) AS TaxFortyTwoOut");
                objStr.Append(", SUM((ISNULL(PAYTRAN_INCOME_401,0) - ISNULL(PAYTRAN_DEDUCT_401,0))) AS IncFortyOne");
                objStr.Append(", SUM((ISNULL(PAYTRAN_INCOME_4013,0) - ISNULL(PAYTRAN_DEDUCT_4013,0))) AS IncFortyOnePerThree");
                objStr.Append(", SUM((ISNULL(PAYTRAN_INCOME_4012,0) - ISNULL(PAYTRAN_DEDUCT_4012,0))) AS IncFortyOneTwo");
                objStr.Append(", SUM((ISNULL(PAYTRAN_INCOME_402I,0) - ISNULL(PAYTRAN_DEDUCT_402I,0))) AS IncFortyTwoIn");
                objStr.Append(", SUM((ISNULL(PAYTRAN_INCOME_402O,0) - ISNULL(PAYTRAN_DEDUCT_402O,0))) AS IncFortyTwoOut");

                objStr.Append(" FROM HRM_TR_PAYTRAN ");
                objStr.Append(" WHERE PAYTRAN_PAYDATE IN (SELECT PERIOD_PAYMENT FROM HRM_MT_PERIOD WHERE YEAR_CODE='" + obj.reportjob_paydate.ToString("yyyy") + "' AND PERIOD_PAYMENT <= '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "' )");
                objStr.Append(" AND COMPANY_CODE = '" + obj.company_code + "' ");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                objStr.Append(" )b ");

                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYTRAN_PAYDATE IN (SELECT PERIOD_PAYMENT FROM HRM_MT_PERIOD WHERE YEAR_CODE='" + obj.reportjob_paydate.ToString("yyyy") + "' AND PERIOD_PAYMENT <= '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "' )");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "tbTRPRAccumulate");
                ds.Tables.Add(dt);

                //tbTRLeaveAccumurate
                //objStr = new StringBuilder();
                //objStr.Append(" SELECT ");
                //objStr.Append(" COMPANY_CODE AS CompID");
                //objStr.Append(", WORKER_CODE AS EmpID");
                //objStr.Append(", YEAR_CODE AS LeaveYear");
                //objStr.Append(", LEAVE_CODE AS LeaveID");
                //objStr.Append(", EMPLEAVEACC_BF AS LeaveBF");
                //objStr.Append(", EMPLEAVEACC_ANNUAL AS AnnualLeave");
                //objStr.Append(", EMPLEAVEACC_USED AS LeaveUsed");

                //objStr.Append(" FROM HRM_TR_EMPLEAVEACC");
                //objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                //objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                //objStr.Append(" AND YEAR_CODE = '" + obj.reportjob_paydate.ToString("yyyy") + "'");

                //dt = objConn.doGetTable(objStr.ToString(), "tbTRLeaveAccumurate");
                //ds.Tables.Add(dt);

                //tbTempReportPNGD
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" a.COMPANY_CODE AS CompID ");
                objStr.Append(", a.WORKER_CODE AS EmpID ");
                objStr.Append(", a.PAYTRAN_PAYDATE AS PayDate ");

                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '01'),0) AS ReducePerson ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '02'),0) AS ReduceSpose ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '06'),0) AS ReduceMather ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '05'),0) AS ReduceFather ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '08'),0) AS ReduceSposeMa ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '07'),0) AS ReduceSposeFa ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '03'),0) AS ReduceChild ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '04'),0) AS ReduceChstudy ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '15'),0) AS ReduceInterance ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = 'PF'),0) AS ReducePF ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = 'RMF'),0) AS ReduceRMF ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = 'SSF'),0) AS ReduceLTF ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '22'),0) AS ReduceInterest ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = 'SSO'),0) AS ReduceSocial ");
                objStr.Append(", (SELECT SUM(EMPREDUCE_AMOUNT) FROM HRM_TR_EMPREDUCE where WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE =  a.COMPANY_CODE) AS TotalReduce ");
                objStr.Append(", 0 AS MoneyPF ");
                objStr.Append(", 0 AS MoneyKBCH ");
                objStr.Append(", 0 AS MoneyTheacher ");
                objStr.Append(", 0 AS MoneyWork ");
                objStr.Append(", 0 AS MoneyAge ");
                objStr.Append(", 0 AS MoneyAgeSpose ");
                objStr.Append(", 0 AS TotalMoney ");
                objStr.Append(", 0 AS DeductForty ");

                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '30'),0) AS DeductStudy ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '31'),0) AS DeductDonation ");
                objStr.Append(", 0 AS TaxOneOne ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '10'),0) AS RInsuFaEmp ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '11'),0) AS RInsuMoEmp ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '12'),0) AS RInsuFaSpo ");
                objStr.Append(", ISNULL((SELECT TOP 1 EMPREDUCE_AMOUNT FROM HRM_TR_EMPREDUCE WHERE WORKER_CODE = a.WORKER_CODE AND COMPANY_CODE = a.COMPANY_CODE AND REDUCE_TYPE = '13'),0) AS RInsuMoSpo ");

                objStr.Append(", 0 AS ReduceTravel ");
                objStr.Append(", 0 AS ReduceCripple ");
                objStr.Append(", 0 AS ReducePremiumPension ");
                objStr.Append(", 0 AS MoneyPropertyValues ");
                objStr.Append(", 0 AS MoneyTaxPropertyValuesDiv5 ");
                objStr.Append(", 0 AS ReduceShoping ");
                objStr.Append(", 0 AS ReduceRealtyPayment ");
                objStr.Append(", 0 AS ReduceRealtyValues ");
                objStr.Append(", 0 AS NationalSavings ");
                objStr.Append(", 0 AS ReduceAntenatalCare ");
                objStr.Append(", 0 AS ReducePartyDonation ");
                objStr.Append(", 0 AS ReduceSSFX ");
                objStr.Append(", 0 AS ReduceSport ");

                objStr.Append(" FROM HRM_TR_PAYTRAN a ");

                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYTRAN_PAYDATE IN (SELECT PERIOD_PAYMENT FROM HRM_MT_PERIOD WHERE YEAR_CODE='" + obj.reportjob_paydate.ToString("yyyy") + "' AND PERIOD_PAYMENT <= '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "' )");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                dt = objConn.doGetTable(objStr.ToString(), "tbTempReportPNGD");
                ds.Tables.Add(dt);

                //tbTRPRTaxResign
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMPANY_CODE AS CompID");
                objStr.Append(", WORKER_CODE AS EmpID");
                objStr.Append(", YEAR(PAYTRAN_PAYDATE) AS PeriodYear");
                objStr.Append(", SUM(0) AS Compensation");
                objStr.Append(", SUM(0) AS OtherIncome");
                objStr.Append(", SUM(0) AS TaxIncome");

                objStr.Append(" FROM HRM_TR_PAYTRAN");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                objStr.Append(" AND PAYTRAN_PAYDATE IN (SELECT PERIOD_PAYMENT FROM HRM_MT_PERIOD WHERE YEAR_CODE='" + obj.reportjob_paydate.ToString("yyyy") + "' AND PERIOD_PAYMENT <= '" + obj.reportjob_paydate.ToString("MM/dd/yyyy") + "' )");
                objStr.Append(" AND WORKER_CODE IN (" + strEmpID + ")");
                objStr.Append(" GROUP BY PAYTRAN_PAYDATE,WORKER_CODE, COMPANY_CODE");
                dt = objConn.doGetTable(objStr.ToString(), "tbTRPRTaxResign");
                ds.Tables.Add(dt);

                //tbMTProvince
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" PROVINCE_CODE AS ProvinceID");
                objStr.Append(", PROVINCE_NAME_EN AS Name");
                objStr.Append(", PROVINCE_NAME_TH AS NameT");
                objStr.Append(" FROM HRM_MT_PROVINCE");
                dt = objConn.doGetTable(objStr.ToString(), "tbMTProvince");
                ds.Tables.Add(dt);

                //tbTRCompBranch
                objStr = new StringBuilder();
                objStr.Append(" SELECT ");
                objStr.Append(" COMBRANCH_CODE AS BranchID");
                objStr.Append(", COMPANY_CODE AS CompID");
                objStr.Append(", COMBRANCH_NAME_TH AS BranchName");
                objStr.Append(", COMBRANCH_NAME_EN AS BranchNameE");
                objStr.Append(", (SELECT TOP 1 PROVINCE_CODE FROM HRM_TR_COMADDRESS WHERE COMPANY_CODE = HRM_TR_COMBRANCH.COMPANY_CODE AND COMBRANCH_CODE = HRM_TR_COMBRANCH.COMBRANCH_CODE) AS ProvinceID");

                objStr.Append(" FROM HRM_TR_COMBRANCH");
                objStr.Append(" WHERE COMPANY_CODE='" + obj.company_code + "'");
                dt = objConn.doGetTable(objStr.ToString(), "tbTRCompBranch");
                ds.Tables.Add(dt);

                //END dt

                strError = "RD.SetDataSource";
                _RD.SetDataSource(ds);

                CrystalReportViewer1.EnableParameterPrompt = false;

                strError = "CrystalReportViewer1.ParameterFieldInfo";
                CrystalReportViewer1.ParameterFieldInfo = paramFields;

                strError = "CrystalReportViewer1.ReportSource";
                CrystalReportViewer1.ReportSource = _RD;

                ds.Dispose();


                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }
        
    }
}