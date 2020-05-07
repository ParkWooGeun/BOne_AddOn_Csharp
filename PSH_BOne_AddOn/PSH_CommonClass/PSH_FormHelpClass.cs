using System;
using CrystalDecisions.CrystalReports.Engine;
using System.Collections.Generic;

using SAPbouiCOM;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn.Form
{
    /// <summary>
    /// Form요소 관련 Helper Class
    /// </summary>
    public class PSH_FormHelpClass
    {
        #region 기본 리포트 구현

        /// <summary>
        /// 크리스탈 리포트 호출 (기본)
        /// </summary>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        public void CrystalReportOpen(string pRptTitle, string pRptName)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            //int loopCount = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                #region ParameterFields 이용(NullReference 예외 발생하여 보류)
                //ParameterFields parameterFields = new ParameterFields();

                //ParameterField[] parameterField = new ParameterField[pDataPack.Count];
                //ParameterDiscreteValue[] parameterRange = new ParameterDiscreteValue[pDataPack.Count];

                //for (int loopCount = 0; loopCount <= pDataPack.Count-1; loopCount ++)
                //{
                //    parameterField[loopCount].ParameterFieldName = pDataPack[loopCount].Code.ToString();
                //    parameterRange[loopCount].Value = pDataPack[loopCount].Value.ToString();
                //    parameterField[loopCount].CurrentValues.Add(parameterRange[loopCount]);
                //}
                //parameterFields.Add(parameterField);
                #endregion

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                //rPT_Viewer1.ReportViewer.ParameterFieldInfo = parameterFields;
                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(100);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트 호출 (Parameter 추가)
        /// </summary>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        public void CrystalReportOpen(string pRptTitle, string pRptName, List<PSH_DataPackClass> pRptParameters)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                for (loopCount = 0; loopCount <= pRptParameters.Count - 1; loopCount++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount].Code.ToString(), pRptParameters[loopCount].Value);
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(100);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트 호출 (Parameter, 비율 추가)
        /// </summary>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pZoomRate">리포트 비율</param>
        public void CrystalReportOpen(string pRptTitle, string pRptName, List<PSH_DataPackClass> pRptParameters, int pZoomRate)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                for (loopCount = 0; loopCount <= pRptParameters.Count - 1; loopCount++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount].Code.ToString(), pRptParameters[loopCount].Value);
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(pZoomRate);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트 호출 (Parameter, Formula 추가)
        /// </summary>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pRptFormulas">리포트로 전달할 Formula</param>
        public void CrystalReportOpen(string pRptTitle, string pRptName, List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pRptFormulas)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;
            int loopCount2 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                for (loopCount1 = 0; loopCount1 <= reportDocument.DataDefinition.FormulaFields.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pRptFormulas.Count - 1; loopCount2++)
                    {
                        if (reportDocument.DataDefinition.FormulaFields[loopCount1].FormulaName == "{" + pRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                        {
                            reportDocument.DataDefinition.FormulaFields[loopCount1].Text = "\"" + pRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                        }
                    }
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(100);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트 호출 (Parameter, Formula, 비율 추가)
        /// </summary>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pRptFormulas">리포트로 전달할 Formula</param>
        /// <param name="pZoomRate">리포트 비율</param>
        public void CrystalReportOpen(string pRptTitle, string pRptName, List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pRptFormulas, int pZoomRate)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;
            int loopCount2 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                for (loopCount1 = 0; loopCount1 <= reportDocument.DataDefinition.FormulaFields.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pRptFormulas.Count - 1; loopCount2++)
                    {
                        if (reportDocument.DataDefinition.FormulaFields[loopCount1].FormulaName == "{" + pRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                        {
                            reportDocument.DataDefinition.FormulaFields[loopCount1].Text = "\"" + pRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                        }
                    }
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(pZoomRate);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        #endregion

        #region 서브 리포트 구현

        /// <summary>
        /// 크리스탈 리포트(서브리포트) 호출 (Parameter, SubReportParameter 추가)
        /// </summary>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pSubRptParameters">SubReport로 전달할 Parameter</param>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        public void CrystalReportOpen(List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pSubRptParameters, string pRptTitle, string pRptName)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                //메인 리포트 파라미터
                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                //Sub 리포트 파라미터
                for (loopCount1 = 0; loopCount1 <= pSubRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pSubRptParameters[loopCount1].Code.ToString(), pSubRptParameters[loopCount1].Value, pSubRptParameters[loopCount1].Type.ToString());
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(100);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트(서브리포트) 호출 (Parameter, Formulas, SubReportParameter 추가)
        /// </summary>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pRptFormulas">리포트로 전달할 Formula</param>
        /// <param name="pSubRptParameters">SubReport로 전달할 Parameter</param>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        public void CrystalReportOpen(List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pRptFormulas, List<PSH_DataPackClass> pSubRptParameters, string pRptTitle, string pRptName)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;
            int loopCount2 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                //메인 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                //메인 리포트 Formula
                for (loopCount1 = 0; loopCount1 <= reportDocument.DataDefinition.FormulaFields.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pRptFormulas.Count - 1; loopCount2++)
                    {
                        if (reportDocument.DataDefinition.FormulaFields[loopCount1].FormulaName == "{" + pRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                        {
                            reportDocument.DataDefinition.FormulaFields[loopCount1].Text = "\"" + pRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                        }
                    }
                }

                //Sub 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pSubRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pSubRptParameters[loopCount1].Code.ToString(), pSubRptParameters[loopCount1].Value, pSubRptParameters[loopCount1].Type.ToString());
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(100);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트(서브리포트) 호출 (Parameter, Formula, SubReportParameter, SubReportFormula 추가)
        /// </summary>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pRptFormulas">리포트로 전달할 Formula</param>
        /// <param name="pSubRptParameters">SubReport로 전달할 Parameter</param>
        /// <param name="pSubRptFormulas">리포트로 전달할 Formula</param>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        public void CrystalReportOpen(List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pRptFormulas, List<PSH_DataPackClass> pSubRptParameters, List<PSH_DataPackClass> pSubRptFormulas, string pRptTitle, string pRptName)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;
            int loopCount2 = 0;
            int loopCount3 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                //메인 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                //메인 리포트 Formula
                for (loopCount1 = 0; loopCount1 <= reportDocument.DataDefinition.FormulaFields.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pRptFormulas.Count - 1; loopCount2++)
                    {
                        if (reportDocument.DataDefinition.FormulaFields[loopCount1].FormulaName == "{" + pRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                        {
                            reportDocument.DataDefinition.FormulaFields[loopCount1].Text = "\"" + pRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                        }
                    }
                }

                //Sub 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pSubRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pSubRptParameters[loopCount1].Code.ToString(), pSubRptParameters[loopCount1].Value, pSubRptParameters[loopCount1].Type.ToString());
                }

                //Sub 리포트 Formula
                for (loopCount1 = 0; loopCount1 <= reportDocument.Subreports.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pSubRptFormulas.Count - 1; loopCount2++)
                    {
                        for (loopCount3 = 0; loopCount3 <= reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields.Count - 1; loopCount3++)
                        {
                            if (reportDocument.Subreports[loopCount1].Name == pSubRptFormulas[loopCount2].Type.ToString())
                            {
                                if (reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields[loopCount3].FormulaName == "{" + pSubRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                                {
                                    reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields[loopCount3].Text = "\"" + pSubRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                                }
                            }
                        }
                    }
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(100);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트(서브리포트) 호출 (Parameter, SubReportParameter, SubReportFormula 추가)
        /// </summary>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pSubRptParameters">SubReport로 전달할 Parameter</param>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        /// <param name="pSubRptFormulas">리포트로 전달할 Formula</param>
        public void CrystalReportOpen(List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pSubRptParameters, string pRptTitle, string pRptName, List<PSH_DataPackClass> pSubRptFormulas)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;
            int loopCount2 = 0;
            int loopCount3 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                //메인 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                //Sub 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pSubRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pSubRptParameters[loopCount1].Code.ToString(), pSubRptParameters[loopCount1].Value, pSubRptParameters[loopCount1].Type.ToString());
                }

                //Sub 리포트 Formula
                for (loopCount1 = 0; loopCount1 <= reportDocument.Subreports.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pSubRptFormulas.Count - 1; loopCount2++)
                    {
                        for (loopCount3 = 0; loopCount3 <= reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields.Count - 1; loopCount3++)
                        {
                            if (reportDocument.Subreports[loopCount1].Name == pSubRptFormulas[loopCount2].Type.ToString())
                            {
                                if (reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields[loopCount3].FormulaName == "{" + pSubRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                                {
                                    reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields[loopCount3].Text = "\"" + pSubRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                                }
                            }
                        }
                    }
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(100);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트(서브리포트) 호출 (Parameter, SubReportParameter, 비율 추가)
        /// </summary>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pSubRptParameters">SubReport로 전달할 Parameter</param>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        /// <param name="pZoomRate">리포트 비율</param>
        public void CrystalReportOpen(List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pSubRptParameters, string pRptTitle, string pRptName, int pZoomRate)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                //메인 리포트 파라미터
                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                //Sub 리포트 파라미터
                for (loopCount1 = 0; loopCount1 <= pSubRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pSubRptParameters[loopCount1].Code.ToString(), pSubRptParameters[loopCount1].Value, pSubRptParameters[loopCount1].Type.ToString());
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(pZoomRate);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트(서브리포트) 호출 (Parameter, Formulas, SubReportParameter 추가)
        /// </summary>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pRptFormulas">리포트로 전달할 Formula</param>
        /// <param name="pSubRptParameters">SubReport로 전달할 Parameter</param>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        /// <param name="pZoomRate">리포트 비율</param>
        public void CrystalReportOpen(List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pRptFormulas, List<PSH_DataPackClass> pSubRptParameters, string pRptTitle, string pRptName, int pZoomRate)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;
            int loopCount2 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                //메인 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                //메인 리포트 Formula
                for (loopCount1 = 0; loopCount1 <= reportDocument.DataDefinition.FormulaFields.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pRptFormulas.Count - 1; loopCount2++)
                    {
                        if (reportDocument.DataDefinition.FormulaFields[loopCount1].FormulaName == "{" + pRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                        {
                            reportDocument.DataDefinition.FormulaFields[loopCount1].Text = "\"" + pRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                        }
                    }
                }

                //Sub 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pSubRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pSubRptParameters[loopCount1].Code.ToString(), pSubRptParameters[loopCount1].Value, pSubRptParameters[loopCount1].Type.ToString());
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(pZoomRate);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트(서브리포트) 호출 (Parameter, Formula, SubReportParameter, SubReportFormula 추가)
        /// </summary>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pRptFormulas">리포트로 전달할 Formula</param>
        /// <param name="pSubRptParameters">SubReport로 전달할 Parameter</param>
        /// <param name="pSubRptFormulas">리포트로 전달할 Formula</param>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        /// <param name="pZoomRate">리포트 비율</param>
        public void CrystalReportOpen(List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pRptFormulas, List<PSH_DataPackClass> pSubRptParameters, List<PSH_DataPackClass> pSubRptFormulas, string pRptTitle, string pRptName, int pZoomRate)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;
            int loopCount2 = 0;
            int loopCount3 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                //메인 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                //메인 리포트 Formula
                for (loopCount1 = 0; loopCount1 <= reportDocument.DataDefinition.FormulaFields.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pRptFormulas.Count - 1; loopCount2++)
                    {
                        if (reportDocument.DataDefinition.FormulaFields[loopCount1].FormulaName == "{" + pRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                        {
                            reportDocument.DataDefinition.FormulaFields[loopCount1].Text = "\"" + pRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                        }
                    }
                }

                //Sub 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pSubRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pSubRptParameters[loopCount1].Code.ToString(), pSubRptParameters[loopCount1].Value, pSubRptParameters[loopCount1].Type.ToString());
                }

                //Sub 리포트 Formula
                for (loopCount1 = 0; loopCount1 <= reportDocument.Subreports.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pSubRptFormulas.Count - 1; loopCount2++)
                    {
                        for (loopCount3 = 0; loopCount3 <= reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields.Count - 1; loopCount3++)
                        {
                            if (reportDocument.Subreports[loopCount1].Name == pSubRptFormulas[loopCount2].Type.ToString())
                            {
                                if (reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields[loopCount3].FormulaName == "{" + pSubRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                                {
                                    reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields[loopCount3].Text = "\"" + pSubRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                                }
                            }
                        }
                    }
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(pZoomRate);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        /// <summary>
        /// 크리스탈 리포트(서브리포트) 호출 (Parameter, SubReportParameter, SubReportFormula 추가)
        /// </summary>
        /// <param name="pRptParameters">리포트로 전달할 Parameter</param>
        /// <param name="pSubRptParameters">SubReport로 전달할 Parameter</param>
        /// <param name="pRptTitle">리포트 제목</param>
        /// <param name="pRptName">리포트 파일(rpt) 명</param>
        /// <param name="pSubRptFormulas">리포트로 전달할 Formula</param>
        /// /// <param name="pZoomRate">리포트 비율</param>
        public void CrystalReportOpen(List<PSH_DataPackClass> pRptParameters, List<PSH_DataPackClass> pSubRptParameters, string pRptTitle, string pRptName, List<PSH_DataPackClass> pSubRptFormulas, int pZoomRate)
        {
            PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1 rPT_Viewer1 = new PSH_BOne_AddOn.EXT_Form.FrmRPT_Viewer1();
            ReportDocument reportDocument = new ReportDocument();

            SAPbouiCOM.ProgressBar ProgBar01 = null;

            int loopCount1 = 0;
            int loopCount2 = 0;
            int loopCount3 = 0;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

                reportDocument.Load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Report + "\\" + pRptName);

                reportDocument.DataSourceConnections[0].IntegratedSecurity = false;
                reportDocument.DataSourceConnections[0].SetConnection(PSH_Globals.SP_ODBC_IP, PSH_Globals.SP_ODBC_DBName, PSH_Globals.SP_ODBC_ID, PSH_Globals.SP_ODBC_PW); //데이터베이스 서버 접속

                //메인 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pRptParameters[loopCount1].Code.ToString(), pRptParameters[loopCount1].Value);
                }

                //Sub 리포트 Parameter
                for (loopCount1 = 0; loopCount1 <= pSubRptParameters.Count - 1; loopCount1++)
                {
                    reportDocument.SetParameterValue(pSubRptParameters[loopCount1].Code.ToString(), pSubRptParameters[loopCount1].Value, pSubRptParameters[loopCount1].Type.ToString());
                }

                //Sub 리포트 Formula
                for (loopCount1 = 0; loopCount1 <= reportDocument.Subreports.Count - 1; loopCount1++)
                {
                    for (loopCount2 = 0; loopCount2 <= pSubRptFormulas.Count - 1; loopCount2++)
                    {
                        for (loopCount3 = 0; loopCount3 <= reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields.Count - 1; loopCount3++)
                        {
                            if (reportDocument.Subreports[loopCount1].Name == pSubRptFormulas[loopCount2].Type.ToString())
                            {
                                if (reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields[loopCount3].FormulaName == "{" + pSubRptFormulas[loopCount2].Code.ToString() + "}") //크리스탈 리포트의 Formula Field(수식 필드)와 DataPack으로 전달한 변수명이 같으면
                                {
                                    reportDocument.Subreports[loopCount1].DataDefinition.FormulaFields[loopCount3].Text = "\"" + pSubRptFormulas[loopCount2].Value.ToString() + "\""; //Formula 변수에 값 저장
                                }
                            }
                        }
                    }
                }

                rPT_Viewer1.ReportViewer.ReportSource = reportDocument;
                rPT_Viewer1.ReportViewer.Refresh();
                rPT_Viewer1.ReportViewer.Zoom(pZoomRate);

                rPT_Viewer1.Text = pRptTitle;

                ProgBar01.Value = 100;
                ProgBar01.Stop();
                ProgBar01 = null;

                rPT_Viewer1.ShowDialog();
            }
            catch (Exception ex)
            {
                ProgBar01.Stop();
                throw ex;
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();

                rPT_Viewer1.ReportViewer.ReportSource = null;
                rPT_Viewer1.Dispose();
                rPT_Viewer1 = null;
            }
        }

        #endregion
    }
}
