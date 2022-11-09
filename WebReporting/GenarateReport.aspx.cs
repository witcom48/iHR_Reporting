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

                        case "PA004":
                            viewerTax("PA004.rpt", objReportjob, objList_whose);
                            break;

                        case "PA005":
                            viewerTax("PA005.rpt", objReportjob, objList_whose);
                            break;

                        case "PA006":
                            viewerTax("PA006.rpt", objReportjob, objList_whose);
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

                if (rptName == "PA006.rpt") 
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
                } else {
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

        
    }
}