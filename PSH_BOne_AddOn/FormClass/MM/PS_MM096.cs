﻿using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 원자재 불출 등록(포장사업팀)
    /// </summary>
    internal class PS_MM096 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_MM096H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM096L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_USERDS01;

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oClick_ColRow;

        private string oBPLId;
        private string oOrdGbn;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM096.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM096_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM096");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                PS_MM096_CreateItems();
                PS_MM096_ComboBox_Setting();
                PS_MM096_Initialization();
                PS_MM096_FormClear();
                PS_MM096_FormItemEnabled();

                oForm.EnableMenu(("1283"), false); // 삭제
                oForm.EnableMenu(("1286"), false); // 닫기
                oForm.EnableMenu(("1287"), false); // 복제
                oForm.EnableMenu(("1284"), true); // 취소
                oForm.EnableMenu(("1293"), true); // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM096_CreateItems()
        {
            try
            {
                oDS_PS_MM096H = oForm.DataSources.DBDataSources.Item("@PS_MM095H");
                oDS_PS_MM096L = oForm.DataSources.DBDataSources.Item("@PS_MM095L");
                oDS_PS_USERDS01 = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;

                oDS_PS_MM096H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM096_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                //작업구분
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where U_PudYN = 'Y' and Code in ('108','109')  Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("OrdGbn").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                //출고구분
                oForm.Items.Item("ExitGbn").Specific.ValidValues.Add("1", "정상출고");
                oForm.Items.Item("ExitGbn").Specific.ValidValues.Add("9", "잔재사용");
                oDS_PS_MM096H.SetValue("U_ExitGbn", 0, "1");

                //작업구분
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("10", "자가");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("30", "외주");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("40", "실적");

                ////작업구분
                oMat02.Columns.Item("WorkGbn").ValidValues.Add("10", "자가");
                oMat02.Columns.Item("WorkGbn").ValidValues.Add("30", "외주");
                oMat02.Columns.Item("WorkGbn").ValidValues.Add("40", "실적");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Initialization
        /// </summary>
        private void PS_MM096_Initialization()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oBPLId))
                {
                    oDS_PS_MM096H.SetValue("U_BPLId", 0, dataHelpClass.User_BPLID());
                }
                else
                {
                    oDS_PS_MM096H.SetValue("U_BPLId", 0, oBPLId); 
                }

                oDS_PS_MM096H.SetValue("U_BPLId", 0, dataHelpClass.User_BPLID());
                oDS_PS_MM096H.SetValue("U_OrdGbn", 0, "108");

                if (oMat01.RowCount == 0)
                {
                    PS_MM096_AddMatrixRow(0, true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM096_HeaderSpaceLineDel()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM096H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM096H.GetValue("U_OrdGbn", 0)))
                {
                    errMessage = "작업구분은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM096H.GetValue("U_DocDate", 0)))
                {
                    errMessage = "전기일은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return functionReturnValue;
        }

        /// <summary>
        /// MatrixSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM096_MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;
            int i;
            string sQry;
            string ItemCode;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                oForm.Freeze(true);
                // 라인
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    //원재료에 대한 제품코드가 누락되었을때
                    if (string.IsNullOrEmpty(oDS_PS_MM096L.GetValue("U_PItemCod", i)) && !string.IsNullOrEmpty(oDS_PS_MM096L.GetValue("U_CItemCod", i)))
                    {
                        errMessage = "제품코드가 없습니다. 확인하세요.";
                        throw new Exception();
                    }

                    //원재료의 불출대상 제품코드가 정상인지 check
                    ItemCode = oDS_PS_MM096L.GetValue("U_PItemCod", i);

                    sQry = "Select Cnt = Count(*) From OITM Where ItmsGrpCod = '102' And ItemCode = '" + ItemCode + "'";

                    oRecordSet01.DoQuery(sQry);

                    if (oRecordSet01.Fields.Item(0).Value == 0)
                    {
                        errMessage = "제품코드를 확인하세요. 원재료 불출대상 제품이 아닙니다.";
                        throw new Exception();
                    }

                }
                oMat01.LoadFromDataSource();

                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    oDS_PS_MM096L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                }
                oMat01.LoadFromDataSource();
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM096_FormItemEnabled()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("IssueYN").Editable = true;
                    oMat01.Columns.Item("IssueQty").Editable = true;
                    oMat01.Columns.Item("IssueWt").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oMat01.Columns.Item("IssueYN").Editable = false;
                    oMat01.Columns.Item("IssueQty").Editable = false;
                    oMat01.Columns.Item("IssueWt").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("OrdGbn").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oMat01.Columns.Item("IssueYN").Editable = false;
                    oMat01.Columns.Item("IssueQty").Editable = false;
                    oMat01.Columns.Item("IssueWt").Editable = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_MM096_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            string WhsCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "Mat01":
                        if (oCol == "PItemCod")
                        {
                            oForm.Freeze(true);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("PItemCod").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_MM096_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }
                            sQry = "Select ItemName  From OITM Where ItemCode = '" + oMat01.Columns.Item("PItemCod").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("PItemNam").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                            oMat01.Columns.Item("DocDate").Cells.Item(oRow).Specific.Value = oForm.Items.Item("DocDate").Specific.Value;
                            oMat01.Columns.Item("WorkGbn").Cells.Item(oRow).Specific.Select("10");

                            oMat01.Columns.Item("PItemCod").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Freeze(false);
                        }
                        else if (oCol == "CItemCod")
                        {
                            oForm.Freeze(true);
                            WhsCode = Convert.ToString(Convert.ToDouble("10") + oForm.Items.Item("BPLId").Specific.Value);
                            sQry = "Select ItemName  From OITM Where ItemCode = '" + oMat01.Columns.Item("CItemCod").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("CItemNam").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                            sQry = "Select Qty = Case When Onhand = 0 Then 0 Else U_Qty End, Weight = Onhand  From OITW";
                            sQry = sQry + " Where ItemCode = '" + oMat01.Columns.Item("CItemCod").Cells.Item(oRow).Specific.Value.ToString().Trim() + "' and WhsCode = '" + WhsCode + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("CQty").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            oMat01.Columns.Item("CWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(1).Value.ToString().Trim();

                            oMat01.Columns.Item("WhsCode").Cells.Item(oRow).Specific.Value = WhsCode;


                            oMat01.Columns.Item("CItemCod").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Freeze(false);
                        }
                        else if (oCol == "WhsCode")
                        {
                            oForm.Freeze(true);
                            WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(oRow).Specific.Value;
                            sQry = "Select Qty = Case When Onhand = 0 Then 0 Else U_Qty End, Weight = Onhand  From OITW";
                            sQry = sQry + " Where ItemCode = '" + oMat01.Columns.Item("CItemCod").Cells.Item(oRow).Specific.Value.ToString().Trim() + "' and WhsCode = '" + WhsCode + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("CQty").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            oMat01.Columns.Item("CWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(1).Value.ToString().Trim();
                            oMat01.Columns.Item("WhsCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Freeze(false);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM096_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM096L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM096L.Offset = oRow;
                oDS_PS_MM096L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_MM096_FormClear()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM095'", "");
                if (Convert.ToDouble(DocNum) == 0)
                {
                    oForm.Items.Item("DocNum").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocNum").Specific.Value = DocNum;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// LoadData
        /// </summary>
        private void PS_MM096_LoadData(int sRow)
        {
            short i = 0;
            string sQry = null;
            string DocEntry = null;
            string BPLID = null;
            string WorkGbn = null;
            string LineId = null;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Items.Item("DocDate").Specific.Value = oMat02.Columns.Item("DocDate").Cells.Item(sRow).Specific.Value.ToString("yyyyMMdd").Trim();

                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                WorkGbn = oMat02.Columns.Item("WorkGbn").Cells.Item(sRow).Specific.Value.ToString().Trim();
                DocEntry = oMat02.Columns.Item("DocEntry").Cells.Item(sRow).Specific.Value.ToString().Trim();
                LineId = oMat02.Columns.Item("LineId").Cells.Item(sRow).Specific.Value.ToString().Trim();

                sQry = "EXEC [PS_MM096_02] '" + BPLID + "','" + WorkGbn + "', '" + DocEntry + "', '" + LineId + "'";
                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_MM096L.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회결과 없습니다. 확인하세요.";
                    throw new Exception();
                }
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
                ProgressBar01.Text = "조회시작!";
                oForm.Freeze(true);
                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_MM096L.Size)
                    {
                        oDS_PS_MM096L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_MM096L.Offset = i;
                    oDS_PS_MM096L.SetValue("U_LineNum", i, Convert.ToString(i + 1));

                    oDS_PS_MM096L.SetValue("U_DocDate", i, oRecordSet01.Fields.Item("DocDate").Value.ToString("yyyyMMdd").Trim());
                    oDS_PS_MM096L.SetValue("U_WorkGbn", i, oRecordSet01.Fields.Item("WorkGbn").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_OrdNum", i, oRecordSet01.Fields.Item("OrdNum").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_CpCode", i, oRecordSet01.Fields.Item("CpCode").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_CpName", i, oRecordSet01.Fields.Item("CpName").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_PItemCod", i, oRecordSet01.Fields.Item("PItemCod").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_PItemNam", i, oRecordSet01.Fields.Item("PItemNam").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_PQty", i, oRecordSet01.Fields.Item("PQty").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_PWeight", i, oRecordSet01.Fields.Item("PWeight").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_CItemCod", i, oRecordSet01.Fields.Item("CItemCod").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_CItemNam", i, oRecordSet01.Fields.Item("CItemNam").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_CutCount", i, oRecordSet01.Fields.Item("CutCount").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_IssueYN", i, oRecordSet01.Fields.Item("IssueYN").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_IssueQty", i, oRecordSet01.Fields.Item("IssueQty").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_IssueWt", i, oRecordSet01.Fields.Item("IssueWt").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_PreQty", i, oRecordSet01.Fields.Item("PreQty").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_PreWt", i, oRecordSet01.Fields.Item("PreWt").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_WhsCode", i, oRecordSet01.Fields.Item("WhsCode").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_CQty", i, oRecordSet01.Fields.Item("CQty").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_CWeight", i, oRecordSet01.Fields.Item("CWeight").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_PP040Doc", i, oRecordSet01.Fields.Item("PP040Doc").Value.ToString().Trim());
                    oDS_PS_MM096L.SetValue("U_PP040Lin", i, oRecordSet01.Fields.Item("PP040Lin").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                }
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// LoadData
        /// </summary>
        private void PS_MM096_LoadData_Mat02()
        {
            short i = 0;
            string sQry = null;
            string BPLID = null;
            string OrdGbn = null;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                OrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();

                sQry = "EXEC [PS_MM096_01] '" + BPLID + "','" + OrdGbn + "'";

                oRecordSet01.DoQuery(sQry);

                oMat02.Clear();

                oDS_PS_USERDS01.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                oForm.Freeze(true);

                ProgressBar01.Text = "조회시작!";

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_USERDS01.Size)
                    {
                        oDS_PS_USERDS01.InsertRecord(i);
                    }

                    oMat02.AddRow();
                    oDS_PS_USERDS01.Offset = i;
                    oDS_PS_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_USERDS01.SetValue("U_ColReg01", i,oRecordSet01.Fields.Item("WorkGbn").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg02", i,oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg03", i,oRecordSet01.Fields.Item("LineId").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg04", i,oRecordSet01.Fields.Item("DocDate").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg05", i,oRecordSet01.Fields.Item("CpName").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg06", i,oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg07", i,oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColNum01", i,oRecordSet01.Fields.Item("Qty").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                }
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
            }
        }

        /// <summary>
        /// LoadData
        /// </summary>
        private void PS_MM096_Update_PP040L(string sStatus)
        {
            string sQry = string.Empty ;
            string WhsCode;
            string CItemCod;
            string IssueYN = string.Empty;
            string PP040Lin;
            string PP040Doc;
            string WorkGbn;
            Double IssueQty;
            Double IssueWt;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (sStatus == "ADD")
                {
                    IssueYN = "Y";
                }
                else if (sStatus == "CANCEL")
                {
                    IssueYN = "N";
                }

                if (oMat01.VisualRowCount > 0)
                {
                    CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(1).Specific.Value;
                    IssueQty = Convert.ToDouble(oMat01.Columns.Item("IssueQty").Cells.Item(1).Specific.Value);
                    IssueWt = Convert.ToDouble(oMat01.Columns.Item("IssueWt").Cells.Item(1).Specific.Value);
                    WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(1).Specific.Value;
                    WorkGbn = oMat01.Columns.Item("WorkGbn").Cells.Item(1).Specific.Value;

                    if (!string.IsNullOrEmpty(CItemCod) && !string.IsNullOrEmpty(WhsCode))
                    {
                        if (oMat01.Columns.Item("IssueYN").Cells.Item(1).Specific.Value.ToString().Trim() == "Y")
                        {
                            PP040Doc = oMat01.Columns.Item("PP040Doc").Cells.Item(1).Specific.Value.ToString().Trim();
                            PP040Lin = oMat01.Columns.Item("PP040Lin").Cells.Item(1).Specific.Value.ToString().Trim();
                            if (WorkGbn == "10")//실동입력
                            {
                                sQry = "Update [@PS_PP040L] Set U_IssueYN = '" + IssueYN + "' Where DocEntry = '" + PP040Doc + "' And LineId = '" + PP040Lin + "'";
                            }
                            else if (WorkGbn == "30")//외주입고
                            {
                                sQry = "Update [@PS_MM138H] Set U_IssueYN = '" + IssueYN + "' Where DocEntry = '" + PP040Doc + "'";
                            }
                            else if (WorkGbn == "40")//생산완료
                            {
                                sQry = "Update [@PS_PP083L] Set U_IssueYN = '" + IssueYN + "' Where DocEntry = '" + PP040Doc + "' And LineId = '" + PP040Lin + "'";
                            }

                            oRecordSet01.DoQuery(sQry);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 출고 DI
        /// </summary>
        /// <returns></returns>
        private bool PS_MM096_Add_oInventoryGenExit()
        {
            bool returnValue = false;
            int i;
            int j = 0;
            int RetVal;
            int errDICode;
            string CItemCod;
            string DocDate;
            string DocNum;
            string WhsCode;
            string errDIMsg;
            string sDocEntry;
            string errMessage = string.Empty;
            double IssueQty;
            double IssueWt;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents DI_oInventoryGenExit = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit); //문서타입(입고)
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();

                PS_MM096_FormClear();
                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errMessage = "현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.";
                    throw new Exception();
                }
                DocNum = oDS_PS_MM096H.GetValue("DocNum", 0).ToString().Trim();
                DocDate = oDS_PS_MM096H.GetValue("U_DocDate", 0);

                if (string.IsNullOrEmpty(oDS_PS_MM096H.GetValue("U_OIGEDoc", 0).ToString().Trim()))
                {
                    PSH_Globals.oCompany.StartTransaction();

                    DI_oInventoryGenExit.DocDate = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                    DI_oInventoryGenExit.TaxDate = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                    DI_oInventoryGenExit.Comments = "원재료 불출 등록(" + DocNum + ") 출고 : PS_MM096";

                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value;
                        IssueQty = Convert.ToDouble(oMat01.Columns.Item("IssueQty").Cells.Item(i + 1).Specific.Value);
                        IssueWt = Convert.ToDouble(oMat01.Columns.Item("IssueWt").Cells.Item(i + 1).Specific.Value);
                        WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i + 1).Specific.Value;

                        if (!string.IsNullOrEmpty(CItemCod) && IssueQty >= 0 && IssueWt != 0 && !string.IsNullOrEmpty(WhsCode))
                        {
                            DI_oInventoryGenExit.Lines.Add();
                            DI_oInventoryGenExit.Lines.SetCurrentLine(j);
                            DI_oInventoryGenExit.Lines.ItemCode = CItemCod;
                            DI_oInventoryGenExit.Lines.WarehouseCode = WhsCode;
                            DI_oInventoryGenExit.Lines.Quantity = IssueWt;
                            DI_oInventoryGenExit.Lines.UserFields.Fields.Item("U_Qty").Value = IssueQty;
                            j++;
                        }
                    }

                    // 완료
                    RetVal = DI_oInventoryGenExit.Add();
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }
                    else
                    {
                        PSH_Globals.oCompany.GetNewObjectCode(out sDocEntry);
                        oDS_PS_MM096H.SetValue("U_OIGEDoc", 0, sDocEntry);
                        PS_MM096_Update_PP040L("ADD");
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                if (DI_oInventoryGenExit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenExit);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM096_InsertoInventoryGenEntry
        /// </summary>
        /// <returns></returns>
        private bool PS_MM096_InsertoInventoryGenEntry()
        {
            bool returnValue = false;
            int i;
            int j;
            int RetVal;
            int errDICode;
            string CItemCod;
            string DocDate;
            string DocNum;
            string WhsCode;
            string errDIMsg;
            string errMessage = string.Empty;
            string sDocEntry;
            string OIGEDoc;
            double IssueQty;
            double IssueWt;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents DI_oInventoryGenEntry = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry); //문서타입(입고)
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();
                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errMessage = "현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.";
                    throw new Exception();
                }
                DocNum = oDS_PS_MM096H.GetValue("DocNum", 0).ToString().Trim();
                DocDate = oDS_PS_MM096H.GetValue("U_DocDate", 0);
                OIGEDoc = oDS_PS_MM096H.GetValue("U_OIGEDoc", 0).ToString().Trim();

                if (string.IsNullOrEmpty(oDS_PS_MM096H.GetValue("U_OIGNDoc", 0).ToString().Trim()))
                {
                    PSH_Globals.oCompany.StartTransaction();

                    DI_oInventoryGenEntry.DocDate = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                    DI_oInventoryGenEntry.TaxDate = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                    DI_oInventoryGenEntry.Comments = "원재료 불출 등록 출고 취소 (" + DocNum + ") 입고 : PS_MM096";
                    DI_oInventoryGenEntry.UserFields.Fields.Item("U_CancDoc").Value = OIGEDoc;

                    j = 0;
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value;
                        IssueQty = Convert.ToDouble(oMat01.Columns.Item("IssueQty").Cells.Item(i + 1).Specific.Value);
                        IssueWt = Convert.ToDouble(oMat01.Columns.Item("IssueWt").Cells.Item(i + 1).Specific.Value);
                        WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i + 1).Specific.Value;

                        if (!string.IsNullOrEmpty(CItemCod) && IssueQty >= 0 && IssueWt != 0 && !string.IsNullOrEmpty(WhsCode))
                        {
                            DI_oInventoryGenEntry.Lines.Add();
                            DI_oInventoryGenEntry.Lines.SetCurrentLine(j);
                            DI_oInventoryGenEntry.Lines.ItemCode = CItemCod;
                            DI_oInventoryGenEntry.Lines.WarehouseCode = WhsCode;
                            DI_oInventoryGenEntry.Lines.Quantity = IssueWt;
                            DI_oInventoryGenEntry.Lines.UserFields.Fields.Item("U_Qty").Value = IssueQty;
                            j += 1;
                        }
                    }

                    // 완료
                    RetVal = DI_oInventoryGenEntry.Add();
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errMessage = "DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg;
                        throw new Exception();
                    }
                    else
                    {
                        PSH_Globals.oCompany.GetNewObjectCode(out sDocEntry);
                        oDS_PS_MM096H.SetValue("U_OIGNDoc", 0, sDocEntry);
                        PS_MM096_Update_PP040L("CANCEL");
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                if (DI_oInventoryGenEntry != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenEntry);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Form Item Event
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">pVal</param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    break;
            }
        }

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        oBPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                        oOrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM096_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_MM096_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (PS_MM096_Add_oInventoryGenExit() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                PS_MM096_Update_PP040L("UPDATE");
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            //if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            //{
                            //    PS_MM096_AddMatrixRow(oMat01.RowCount, false);
                            //    oLast_Mode = 100;
                            //}
                            //else if (oLast_Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            //{
                            //    PS_MM096_AddMatrixRow(oMat01.RowCount, false);
                            //    PS_MM096_FormItemEnabled();
                            //    oLast_Mode = 100;
                            //}
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            oMat02.Clear();
                            oDS_PS_USERDS01.Clear();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "PItemCod")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("PItemCod").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ColUID == "CItemCod")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("CItemCod").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat02")
                    {

                        if (pVal.ColUID == "ItemCode")
                        {
                            if (!string.IsNullOrEmpty(oMat02.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.String))
                            {
                                PS_MM002 PS_MM002 = new PS_MM002();
                                PS_MM002.LoadForm(oForm.Items.Item("BPLId").Specific.Value, oMat02.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.String);
                                BubbleEvent = false;
                            }
                            else
                            {

                            }
                        }
                        else if (pVal.ColUID == "DocEntry")
                        {
                            if (oMat02.Columns.Item("WorkGbn").Cells.Item(pVal.Row).Specific.Value == "10")
                            {
                                PS_PP043 PS_PP043 = new PS_PP043();
                                PS_PP043.LoadForm(oMat02.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.String);
                                BubbleEvent = false;
                            }
                            //else if (oMat02.Columns.Item("WorkGbn").Cells.Item(pVal.Row).Specific.Value == "30")
                            //{
                            //    PS_MM138 PS_MM138 = new PS_MM138();
                            //    PS_MM138.LoadForm(oMat02.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.String);
                            //    BubbleEvent = false;
                            //}
                            //else if (oMat02.Columns.Item("WorkGbn").Cells.Item(pVal.Row).Specific.Value == "40")
                            //{
                            //    PS_PP083 PS_PP083 = new PS_PP083();
                            //    PS_PP083.LoadForm(oMat02.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.String);
                            //    BubbleEvent = false;

                            //}
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID01 = pVal.ItemUID;
                    oLastColUID01 = "";
                    oLastColRow01 = 0;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "BPLId" || pVal.ItemUID == "OrdGbn")
                        {
                            PS_MM096_LoadData_Mat02();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat02" && pVal.ColUID == "LineNum" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oClick_ColRow = pVal.Row;
                        PS_MM096_LoadData(oClick_ColRow);
                    }
                }
            } 
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string WhsCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "PItemCod")
                            {
                                PS_MM096_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "CItemCod")
                            {
                                PS_MM096_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "WhsCode")
                            {
                                oForm.Freeze(true);
                                oMat01.FlushToDataSource();
                                WhsCode = oDS_PS_MM096L.GetValue("U_WhsCode", pVal.Row - 1).ToString().Trim();
                                if (string.IsNullOrEmpty(WhsCode))
                                {
                                    oDS_PS_MM096L.SetValue("U_CQty", pVal.Row - 1, Convert.ToString(0));
                                    oDS_PS_MM096L.SetValue("U_CWeight", pVal.Row - 1, Convert.ToString(0));
                                    oMat01.LoadFromDataSource();
                                    oMat01.Columns.Item("WhsCode").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    dataHelpClass.MDC_GF_Message("창고는 필수 선택입니다.", "E");
                                }
                                else
                                {

                                    PS_MM096_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                                }
                                oForm.Freeze(false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FORM_UNLOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    SubMain.Remove_Forms(oFormUniqueID);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oForm.Items.Item("Mat02").Top = 81;
                    oForm.Items.Item("Mat02").Left = 6;
                    oForm.Items.Item("Mat02").Width = oForm.Width - 18;
                    oForm.Items.Item("Mat02").Height = oForm.Height / 3;
                    oForm.Items.Item("Mat01").Top = oForm.Items.Item("Mat02").Top + oForm.Height / 3;
                    oForm.Items.Item("Mat01").Left = 6;
                    oForm.Items.Item("Mat01").Width = oForm.Width - 18;
                    oForm.Items.Item("Mat01").Height = Convert.ToInt32(Convert.ToDouble(oForm.Height) / 2.4);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            if (oDS_PS_MM096H.GetValue("Canceled", 0).ToString().Trim() == "N")
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("해당 문서가 취소됩니다 계속하시겠습니까??", 1, "&확인", "&취소") != 1)
                                {
                                    BubbleEvent = false;
                                    return;
                                }

                                if (PS_MM096_InsertoInventoryGenEntry() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                dataHelpClass.MDC_GF_Message("취소된 생산실적입니다. 확인하세요.", "E");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            PS_MM096_FormItemEnabled();
                            //oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PS_MM096L.RemoveRecord(oDS_PS_MM096L.Size - 1);
                                // Mat01에 마지막라인(빈라인) 삭제
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();
                            }
                            break;
                        case "1281": //찾기
                            PS_MM096_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            oForm.Freeze(true);
                            PS_MM096_Initialization();
                            PS_MM096_FormItemEnabled();
                            PS_MM096_FormClear();

                            if (string.IsNullOrEmpty(oBPLId))
                            {
                                oForm.Items.Item("BPLId").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                            else
                            {
                                oForm.Items.Item("BPLId").Specific.Select(oBPLId, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }

                            if (string.IsNullOrEmpty(oOrdGbn))
                            {
                                oForm.Items.Item("OrdGbn").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                            else
                            {
                                oForm.Items.Item("OrdGbn").Specific.Select(oOrdGbn, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }

                            oForm.Items.Item("ExitGbn").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

                            oDS_PS_MM096H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            oForm.Freeze(false);
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            PS_MM096_FormItemEnabled();
                            break;
                        case "1287": //복제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                }

                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }
    }
}
