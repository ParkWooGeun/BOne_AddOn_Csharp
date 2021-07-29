using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 작업지시리스트조회
    /// </summary>
    internal class PS_PP035 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP035L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP035.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP035_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP035");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_PP035_CreateItems();
                PS_PP035_ComboBox_Setting();
                PS_PP035_CF_ChooseFromList();
                PS_PP035_EnableMenus();
                PS_PP035_SetDocument(oFormDocEntry);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PS_PP035_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oDS_PS_PP035L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.DataSources.UserDataSources.Add("Canceled", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Canceled").Specific.DataBind.SetBound(true, "", "Canceled");

                oForm.DataSources.UserDataSources.Add("OrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("OrdNum").Specific.DataBind.SetBound(true, "", "OrdNum");

                oForm.DataSources.UserDataSources.Add("OrdNum1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("OrdNum1").Specific.DataBind.SetBound(true, "", "OrdNum1");

                oForm.DataSources.UserDataSources.Add("OrdGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("OrdGbn").Specific.DataBind.SetBound(true, "", "OrdGbn");

                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                oForm.DataSources.UserDataSources.Add("WorkDtFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("WorkDtFr").Specific.DataBind.SetBound(true, "", "WorkDtFr");

                oForm.DataSources.UserDataSources.Add("WorkDtTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("WorkDtTo").Specific.DataBind.SetBound(true, "", "WorkDtTo");

                oForm.DataSources.UserDataSources.Add("FrgnName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("FrgnName").Specific.DataBind.SetBound(true, "", "FrgnName");

                oForm.DataSources.UserDataSources.Add("Size", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Size").Specific.DataBind.SetBound(true, "", "Size");

                oForm.DataSources.UserDataSources.Add("PrdYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("PrdYN").Specific.DataBind.SetBound(true, "", "PrdYN");

                oForm.DataSources.UserDataSources.Add("ChkWCon", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("ChkWCon").Specific.DataBind.SetBound(true, "", "ChkWCon");
                oForm.Items.Item("Mat01").Enabled = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP035_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                // 사업장
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId",  "1",  false,  false);

                // 작지상태
                dataHelpClass.Combo_ValidValues_Insert("PS_PP035", "Canceled", "", "N", "계획");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP035", "Canceled", "", "Y", "취소");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("Canceled").Specific, "PS_PP035", "Canceled", false);
                oForm.Items.Item("Canceled").Specific.Select("계획", SAPbouiCOM.BoSearchKey.psk_ByValue);

                // 작지구분
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y'", "101", false, false);

                if (dataHelpClass.User_BPLID() == "1")
                {
                    oForm.Items.Item("OrdGbn").Specific.Select("104", SAPbouiCOM.BoSearchKey.psk_ByDescription);
                }
                else
                {
                    oForm.Items.Item("OrdGbn").Specific.Select("105", SAPbouiCOM.BoSearchKey.psk_ByDescription);
                }

                //작지상태(matrix)
                dataHelpClass.Combo_ValidValues_Insert("PS_PP035", "oMat01", "Canceled", "N", "계획");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP035", "oMat01", "Canceled", "Y", "취소");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("Canceled"), "PS_PP035", "oMat01", "Canceled", false);
                dataHelpClass.Set_ComboList(oForm.Items.Item("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] order by Code", "", false, true);

                //작업상태(2012.06.30 송명규 추가)
                sQry = "SELECT T1.U_Minor,T1.U_CdName";
                sQry += " FROM [@PS_SY001H] AS T0 INNER JOIN [@PS_SY001L] AS T1 ON T0.Code = T1.Code";
                sQry += " WHERE T1.Code = 'S003'";

                oMat01.Columns.Item("WorkCon").ValidValues.Add("", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("WorkCon"), sQry,"","");
                
                //작업상태(2012.06.30 송명규 추가)
                oForm.Items.Item("TradeType").Specific.ValidValues.Add("1", "전체");
                oForm.Items.Item("TradeType").Specific.ValidValues.Add("2", "일반");
                oForm.Items.Item("TradeType").Specific.ValidValues.Add("3", "선생산");
                oForm.Items.Item("TradeType").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

                //생산완료 구분 추가(2011.11.29 송명규)
                oForm.Items.Item("PrdYN").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("PrdYN").Specific.ValidValues.Add("Y", "생산완료");
                oForm.Items.Item("PrdYN").Specific.ValidValues.Add("N", "생산미완료");
                oForm.Items.Item("PrdYN").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PS_PP035_CF_ChooseFromList()
        {
            SAPbouiCOM.EditText oEdit = null; 
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            try
            {
                oEdit = oForm.Items.Item("ItemCode").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLITEMCD";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oEdit.ChooseFromListUID = "CFLITEMCD";
                oEdit.ChooseFromListAlias = "ItemCode";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if(oCFLs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                }
                if (oEdit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
                }
                if (oCFL != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                }
                if (oCFLCreationParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                }
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP035_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1281", false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_PP035_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP035_AddMatrixRow(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Chk_Data
        /// </summary>
        [STAThread]
        private void Chk_Data()
        {
            int i;
            string WinTitle;
            string ReportName;
            string sQry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "DELETE Temp_LBG10";
                oRecordSet01.DoQuery(sQry);

                // 조회화면에서 선택한 문서번호만 임시테이블에 삽입
                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    oDS_PS_PP035L.Offset = i;
                    if (oDS_PS_PP035L.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                    {
                        sQry = "INSERT INTO Temp_LBG10 (DocEntry) VALUES  (" + oDS_PS_PP035L.GetValue("U_ColReg02", i).ToString().Trim() + ")";
                        oRecordSet01.DoQuery(sQry);
                    }
                }
                WinTitle = "[PS_PP035] 자주검사 CHECK SHEET";
                ReportName = "PS_PP035_03.rpt";

                formHelpClass.CrystalReportOpen(WinTitle, ReportName);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Chk_Data1
        /// </summary>
        [STAThread]
        private void Chk_Data1()
        {
            int i; 
            string WinTitle;
            string ReportName;
            string sQry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>();

                if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "105" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "106")
                {
                    WinTitle = "[PS_PP034] 공정작업지시서";
                    ReportName = "PS_PP035_04.rpt";

                    //문서번호 입력된 임시테이블 삭제
                    sQry = "DELETE Temp_LBG11";
                    oRecordSet01.DoQuery(sQry);

                    //조회화면에서 선택한 문서번호만 임시테이블에 삽입
                    oMat01.FlushToDataSource();
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        oDS_PS_PP035L.Offset = i;
                        if (oDS_PS_PP035L.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                        {
                            sQry = "INSERT INTO Temp_LBG11 (DocEntry) VALUES  (" + oDS_PS_PP035L.GetValue("U_ColReg02", i).ToString().Trim() + ")";
                            oRecordSet01.DoQuery(sQry);
                        }
                    }
                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@Seq", "M"));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@Seq", "S", "Sub_Report01"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@Seq", "E", "Sub_Report02"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackSubReportParameter, WinTitle, ReportName);
                }
                else if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "101" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "102")
                {
                    WinTitle = "[PS_PP035] 작업지시서";
                    ReportName = "PS_PP035_05.rpt"; 
                    
                    sQry = "DELETE Temp_LBG12";
                    oRecordSet01.DoQuery(sQry);

                    // 조회화면에서 선택한 문서번호만 임시테이블에 삽입
                    oMat01.FlushToDataSource();
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        oDS_PS_PP035L.Offset = i;
                        if (oDS_PS_PP035L.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                        {
                            sQry = "INSERT INTO Temp_LBG12 (DocEntry) VALUES  (" + oDS_PS_PP035L.GetValue("U_ColReg02", i).ToString().Trim() + ")";
                            oRecordSet01.DoQuery(sQry);
                        }
                    }
                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@Seq", "M"));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@Seq", "S", "Sub_Report01"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackSubReportParameter, WinTitle, ReportName);
                }
                else if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "107")
                {
                    WinTitle = "[PS_PP035] End Bearing  공정카드";
                    ReportName = "PS_PP035_06.rpt";

                    sQry = "DELETE Temp_LBG13";
                    oRecordSet01.DoQuery(sQry);

                    // 조회화면에서 선택한 문서번호만 임시테이블에 삽입
                    oMat01.FlushToDataSource();
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        oDS_PS_PP035L.Offset = i;
                        if (oDS_PS_PP035L.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                        {
                            sQry = "INSERT INTO Temp_LBG13 (DocEntry) VALUES  (" + oDS_PS_PP035L.GetValue("U_ColReg02", i).ToString().Trim() + ")";
                            oRecordSet01.DoQuery(sQry);
                        }
                    }
                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@Seq", "M"));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@Seq", "S", "Sub_Report01"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackSubReportParameter, WinTitle, ReportName);
                }
                else if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "104")
                {
                    WinTitle = "[PS_PP035] M/G 공정카드";
                    ReportName = "PS_PP035_07.rpt";

                    sQry = "DELETE Temp_LBG14";
                    oRecordSet01.DoQuery(sQry);

                    // 조회화면에서 선택한 문서번호만 임시테이블에 삽입
                    oMat01.FlushToDataSource();
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    { 
                        oDS_PS_PP035L.Offset = i;
                        if (oDS_PS_PP035L.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
                        {
                            sQry = "INSERT INTO Temp_LBG14 (DocEntry) VALUES  (" + oDS_PS_PP035L.GetValue("U_ColReg02", i).ToString().Trim() + ")";
                            oRecordSet01.DoQuery(sQry);
                        }
                    }
                    formHelpClass.CrystalReportOpen(WinTitle, ReportName);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PS_PP035_SetBaseForm
        /// </summary>
        private void PS_PP035_SetBaseForm()
        {
            int i;
            string Query01;
            string Param01;
            string Param02;
            double Param03;
            string Param04;
            string Param05;
            string Param06;
            string Param07;
            string ItemCode;
            string ItemName;
            string WorkCon;
            string QueryString;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                for (i = 1; i <= oMat01.RowCount; i++)
                {
                    if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
                    {
                        if (oForm.Items.Item("ChkWCon").Specific.Checked == true)//작업상태만 변경 시 재고존재 체크 안함
                        {
                            Param01 = oMat01.Columns.Item("DocEntry").Cells.Item(i).Specific.Value;
                            WorkCon = oMat01.Columns.Item("WorkCon").Cells.Item(i).Specific.Value;

                            Query01 = "EXEC PS_PP035_80 '" + Param01 + "', '" + WorkCon + "'";
                            oRecordSet01.DoQuery(Query01);
                            PSH_Globals.SBO_Application.MessageBox("작업상태를 수정하였습니다.");
                        }
                        else
                        {
                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value.ToString().Trim() == "104")
                            {
                                QueryString =  " SELECT Sum(a.Quantity * (Case When a.Direction = '0' Then 1 Else -1 End)) As Quantity";
                                QueryString += "   FROM IBT1 a Inner Join  OITM b  On a.ItemCode = b.ItemCode And b.U_ItmBsort = '104'";
                                QueryString += "  WHERE a.BaseType In ('59', '60')";
                                QueryString += "    AND a.BatchNum = '" + oMat01.Columns.Item("OrdNum").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                            }
                            else
                            {
                                QueryString =  "SELECT SUM(A.InQty) - SUM(A.OutQty) AS [StockQty]";
                                QueryString += "  FROM OINM AS A INNER JOIN OITM As B ON A.ItemCode = B.ItemCode";
                                QueryString += "  WHERE B.U_ItmBsort IN ('105','106')";
                                QueryString += "    AND A.ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                QueryString += "  GROUP BY  A.ItemCode";
                            }
                            if ((string.IsNullOrEmpty(dataHelpClass.GetValue(QueryString, 0, 1)) ? 0 : Convert.ToInt32(dataHelpClass.GetValue(QueryString, 0, 1))) > 0)
                            {
                                errMessage = i + "행의 작업지시는 현재 재고가 존재합니다. 처리할 수 없습니다.";
                                throw new Exception();
                            }
                            else
                            {
                                Param01 = oMat01.Columns.Item("DocEntry").Cells.Item(i).Specific.Value;
                                Param02 = oMat01.Columns.Item("Canceled").Cells.Item(i).Specific.Value;
                                Param03 = Convert.ToDouble(oMat01.Columns.Item("SelWt").Cells.Item(i).Specific.Value);
                                Param04 = oMat01.Columns.Item("CntcCode").Cells.Item(i).Specific.Value;
                                Param05 = oMat01.Columns.Item("CntcName").Cells.Item(i).Specific.Value;
                                Param06 = oMat01.Columns.Item("DocDate").Cells.Item(i).Specific.Value;
                                Param07 = oMat01.Columns.Item("DueDate").Cells.Item(i).Specific.Value;
                                ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
                                ItemName = dataHelpClass.Make_ItemName(oMat01.Columns.Item("ItemName").Cells.Item(i).Specific.Value.ToString().Trim());

                                Query01 = "EXEC PS_PP035_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "', '" + ItemCode + "', '" + ItemName + "'";
                                oRecordSet01.DoQuery(Query01);
                                PSH_Globals.SBO_Application.MessageBox("데이터를 수정하였습니다.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PS_PP035_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP035_AddMatrixRow(int oRow, bool RowIserted = false)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false) //행추가여부
                {
                    oDS_PS_PP035L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP035L.Offset = oRow;
                oDS_PS_PP035L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_PP035_MTX01
        /// </summary>
        private void PS_PP035_MTX01()
        {
            int i;
            string Query01;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            string Param07;
            string Param08;
            string Param09;
            string Param10;
            string Param11;
            string Param12;
            string Param13;
            string Param14;
            string Temp01;
            string Temp02;
            string Temp03;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(); ;
                Param02 = oForm.Items.Item("Canceled").Specific.Value.ToString().Trim();
                Param03 = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
                Param04 = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
                Param05 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                Param06 = oForm.Items.Item("WorkDtFr").Specific.Value.ToString().Trim();
                Param07 = oForm.Items.Item("WorkDtTo").Specific.Value.ToString().Trim();
                Param08 = oForm.Items.Item("FrgnName").Specific.Value.ToString().Trim();
                Param09 = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
                Param10 = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                Param11 = oForm.Items.Item("Mark").Specific.Value.ToString().Trim();
                Param12 = oForm.Items.Item("OrdNum1").Specific.Value.ToString().Trim();
                Param13 = oForm.Items.Item("TradeType").Specific.Value.ToString().Trim();
                Param14 = oForm.Items.Item("PrdYN").Specific.Value.ToString().Trim();

                if((Param03 + Param12 == "") && (Param06 + Param07 == ""))
                {
                    errMessage = "작지관리번호나 지시일자 필수입니다.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(Param06) && string.IsNullOrEmpty(Param07))
                {
                    Param06 = "19000101";
                    Param07 = "99991231";
                }
                else if (string.IsNullOrEmpty(Param06) && !string.IsNullOrEmpty(Param07))
                {
                    Param06 = "";
                }
                else if (!string.IsNullOrEmpty(Param06) && string.IsNullOrEmpty(Param07))
                {
                    Param07 = Param06;
                }
                else if (!string.IsNullOrEmpty(Param06) && !string.IsNullOrEmpty(Param07))
                {
                }
                if (string.IsNullOrEmpty(Param03) && string.IsNullOrEmpty(Param12))
                {
                    Param03 = "0";
                    Param12 = "ZZZZZZZZZZZZZZZZZZZ";
                }
                else if (string.IsNullOrEmpty(Param03))
                {
                    Param03 = Param12;
                }
                else if (string.IsNullOrEmpty(Param12))
                {
                    Param12 = Param03;
                }
                ProgressBar01.Text = "조회시작!";
                Query01 = "EXEC PS_PP035_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "', '" + Param08 + "', '" + Param09 + "', '" + Param10 + "', '" + Param11 + "', '" + Param12 + "', '" + Param13 + "','" + Param14 + "'";
                oRecordSet01.DoQuery(Query01);
                
                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    oForm.Items.Item("Mat01").Enabled = false;
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                else
                {
                    oForm.Items.Item("Mat01").Enabled = true;
                }
                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_PP035L.InsertRecord(i);
                    }
                    oDS_PS_PP035L.Offset = i;
                    oDS_PS_PP035L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP035L.SetValue("U_ColReg01", i, Convert.ToString(false));
                    oDS_PS_PP035L.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("DocEntry").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("OrdGbn").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("OrdNum").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("OrdSub1").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("OrdSub2").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("Canceled").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("CardCode").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("CardName").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("ItemCode").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("ItemName").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("Quality").Value);
                    oDS_PS_PP035L.SetValue("U_ColNum01", i, oRecordSet01.Fields.Item("Unweight").Value);
                    oDS_PS_PP035L.SetValue("U_ColNum02", i, oRecordSet01.Fields.Item("ReqWt").Value);
                    oDS_PS_PP035L.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("SelWt").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg16", i, oRecordSet01.Fields.Item("CntcCode").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("CntcName").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg17", i, oRecordSet01.Fields.Item("SubName").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg18", i, oRecordSet01.Fields.Item("JISNO").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg19", i, oRecordSet01.Fields.Item("BatchNum").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg25", i, oRecordSet01.Fields.Item("OutSize").Value);
                    oDS_PS_PP035L.SetValue("U_ColReg20", i, oRecordSet01.Fields.Item("WorkCon").Value);
                    //작업상태(2012.06.30 송명규 추가)
                    if (oRecordSet01.Fields.Item("DocDate").Value.ToString("yyyyMMdd").Trim() == "19000101")
                    {
                        Temp01 = "";
                    }
                    else
                    {
                        Temp01 = oRecordSet01.Fields.Item("DocDate").Value.ToString().Trim();
                    }
                    Temp01 = DateTime.Parse(Temp01).ToString("yyyyMMdd");
                    oDS_PS_PP035L.SetValue("U_ColDt01", i, Temp01);

                    if (oRecordSet01.Fields.Item("DueDate").Value.ToString("yyyyMMdd").Trim() == "19000101")
                    {
                        Temp02 = "";
                    }
                    else
                    {
                        Temp02 = oRecordSet01.Fields.Item("DueDate").Value.ToString().Trim();
                    }
                    Temp02 = DateTime.Parse(Temp02).ToString("yyyyMMdd");
                    
                    oDS_PS_PP035L.SetValue("U_ColDt02", i, Temp02);

                    if (oRecordSet01.Fields.Item("WalDoc").Value.ToString("yyyyMMdd").Trim() == "19000101")
                    {
                        Temp03 = "";
                    }
                    else
                    {
                        Temp03 = oRecordSet01.Fields.Item("WalDoc").Value.ToString().Trim();
                    }
                    Temp03 = DateTime.Parse(Temp03).ToString("yyyyMMdd");
                   
                    oDS_PS_PP035L.SetValue("U_ColDt03", i, Temp03);
                    oDS_PS_PP035L.SetValue("U_ColReg15", i, oRecordSet01.Fields.Item("CpName").Value);

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                if (Param04 == "104")
                {
                    oMat01.Columns.Item("ItemCode").Editable = true;
                }
                else
                {
                    oMat01.Columns.Item("ItemCode").Editable = false;
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                dataHelpClass.MDC_GF_Message("조회를 성공하엿습니다.", "S");
            }
            catch (Exception ex)
            {
                if (errMessage != null)
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
                oForm.Update();
                oForm.Freeze(false); ;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP035_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP035'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
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

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                    if (pVal.ItemUID == "Btn01")
                    {
                        PS_PP035_MTX01();
                    }
                    if (pVal.ItemUID == "Btn02")
                    {
                       PS_PP035_SetBaseForm();
                    }
                    if (pVal.ItemUID == "Btn03")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(Chk_Data1);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    if (pVal.ItemUID == "Btn04")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(Chk_Data);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CntcCode");

                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "ItemCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02")
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string Query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "DocEntry")
                            {
                            }
                            else if (pVal.ColUID == "CntcCode")
                            {
                                Query01 = "SELECT LastName, FirstName FROM [OHEM] WHERE EmpID = '" + oMat01.Columns.Item("CntcCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(Query01);
                                oMat01.Columns.Item("CntcName").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim() + oRecordSet01.Fields.Item(1).Value.ToString().Trim();
                            }
                            else if (pVal.ColUID == "ItemCode")
                            {
                                Query01 = "Select ItemName From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(Query01);
                                oMat01.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            }
                            else
                            {
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Update();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.String))
                    {
                        PS_PP030 PS_PP030 = new PS_PP030();
                        PS_PP030.LoadForm(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.String);
                        BubbleEvent = false;
                    }
                    else
                    {
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP035L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "ItemCode")
                    {
                        if (oDataTable01 == null)
                        {
                        }
                        else
                        {
                            oForm.DataSources.UserDataSources.Item("ItemCode").Value = oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value;
                            oForm.DataSources.UserDataSources.Item("ItemName").Value = oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value;
                        }
                    }
                    oForm.Update();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
            }
        }

        /// <summary>
        /// Raise_EVENT_DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string Chk = string.Empty;
            
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" & pVal.Row == Convert.ToDouble("0") & pVal.ColUID == "CHK")
                    {
                        oMat01.FlushToDataSource();
                        if (string.IsNullOrEmpty(oDS_PS_PP035L.GetValue("U_ColReg01", 0).ToString().Trim()) || oDS_PS_PP035L.GetValue("U_ColReg01", 0).ToString().Trim() == "N")
                        {
                            Chk = "Y";
                        }
                        else if (oDS_PS_PP035L.GetValue("U_ColReg01", 0).ToString().Trim() == "Y")
                        {
                            Chk = "N";
                        }
                        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            oDS_PS_PP035L.SetValue("U_ColReg01", i, Chk);
                        }
                        oMat01.LoadFromDataSource();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
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
                        case "7169":
                            PS_PP035_AddMatrixRow(oMat01.VisualRowCount);//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                        case "7169":
                            oDS_PS_PP035L.RemoveRecord(oDS_PS_PP035L.Size - 1); //엑셀 내보내기 이후 처리
                            oMat01.LoadFromDataSource();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02")
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
