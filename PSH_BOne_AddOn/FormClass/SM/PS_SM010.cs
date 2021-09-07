using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// File           : PS_SM010.cls
    /// Module         : 재고관리 > 재고장
    /// Desc           : 재고장
    /// FormType       : PS_SM010
    /// Create Date    : 2010.08.20
    /// Copyright  (c) Morning Data
    /// </summary>
    internal class PS_SM010 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Grid oGrid01;
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private SAPbouiCOM.Form oBaseForm01;
        private string oBaseItemUID01;
        private string oBaseColUID01;
        private int oBaseColRow01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            this.LoadForm();
        }

        /// <summary>
        /// Form 호출(다른 폼에서 호출)
        /// </summary>
        /// <param name="baseForm">기준 Form</param>
        /// <param name="baseItemUID">기준 Form의 ItemUID</param>
        /// <param name="baseColUID">기준 Form의 Matrix ColUID</param>
        /// <param name="baseMatRow">기준 Form의 Matrix Row</param>
        public void LoadForm(SAPbouiCOM.Form baseForm, string baseItemUID, string baseColUID, int baseMatRow)
        {
            oBaseForm01 = baseForm;
            oBaseItemUID01 = baseItemUID;
            oBaseColUID01 = baseColUID;
            oBaseColRow01 = baseMatRow;

            this.LoadForm();
        }

        /// <summary>
        /// Form 호출
        /// </summary>
        private void LoadForm()
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SM010.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_SM010_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SM010");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_SM010_CreateItems();
                PS_SM010_ComboBox_Setting();
                PS_SM010_FormItemEnabled();

                oForm.EnableMenu("1283", true); //삭제
                oForm.EnableMenu("1287", true); //복제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
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
        private void PS_SM010_CreateItems()
        {
            try
            {
                oGrid01 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                oForm.DataSources.UserDataSources.Add("StockType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("StockType").Specific.DataBind.SetBound(true, "", "StockType");

                oForm.DataSources.UserDataSources.Add("TradeType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TradeType").Specific.DataBind.SetBound(true, "", "TradeType");

                oForm.DataSources.UserDataSources.Add("ItemGpCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItemGpCd").Specific.DataBind.SetBound(true, "", "ItemGpCd");

                oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

                oForm.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");

                oForm.DataSources.UserDataSources.Add("Size", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Size").Specific.DataBind.SetBound(true, "", "Size");

                oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

                oForm.DataSources.UserDataSources.Add("Mark", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Mark").Specific.DataBind.SetBound(true, "", "Mark");

                oForm.DataSources.UserDataSources.Add("Check01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

                oForm.Items.Item("ChkBox01").Specific.ValOn = "Y";
                oForm.Items.Item("ChkBox01").Specific.ValOff = "N";
                oForm.Items.Item("ChkBox01").Specific.DataBind.SetBound(true, "", "Check01");
                oForm.DataSources.UserDataSources.Item("Check01").Value = "N";
                //미체크로 값을 주고 폼을 로드
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SM010_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                dataHelpClass.Combo_ValidValues_Insert("PS_SM010", "StockType", "", "1", "재고있는품목");
                dataHelpClass.Combo_ValidValues_Insert("PS_SM010", "StockType", "", "2", "전체");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("StockType").Specific, "PS_SM010", "StockType", false);
                oForm.Items.Item("StockType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);

                dataHelpClass.Combo_ValidValues_Insert("PS_SM010", "TradeType", "", "", "전체");
                dataHelpClass.Combo_ValidValues_Insert("PS_SM010", "TradeType", "", "1", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_SM010", "TradeType", "", "2", "임가공");
                dataHelpClass.Combo_ValidValues_SetValueItem((oForm.Items.Item("TradeType").Specific), "PS_SM010", "TradeType", false);
                oForm.Items.Item("TradeType").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //oForm.Items.Item("ItmBsort").Specific.ValidValues.Add("선택", "선택");
                //dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", false, false);

                //oForm.Items.Item("ItmMsort").Specific.ValidValues.Add("선택", "선택");
                //dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", "", false, false);

                oForm.Items.Item("ItemType").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", "", false, false);

                oForm.Items.Item("Mark").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", "", false, false);

                oForm.Items.Item("ItemGpCd").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemGpCd").Specific, "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]", "", false, false);

                //oForm.Items.Item("ItmBsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //oForm.Items.Item("ItmMsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Mark").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("ItemGpCd").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 창고
                sQry = "SELECT WhsCode, WhsName From OWHS";
                oRecordSet01.DoQuery(sQry);

                oForm.Items.Item("WhsCode").Specific.ValidValues.Add("선택", "선택");

                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("WhsCode").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("WhsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.Items.Item("TradeType").Enabled = false;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_SM010_FormItemEnabled()
        {
            bool itemBSortYN = false;

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    for (int i = 0; i < oBaseForm01.Items.Count; i++) //BaseForm에 ItmBSort가 있는 검사
                    {
                        if (oBaseForm01.Items.Item(i).UniqueID == "ItmBSort")
                        {
                            itemBSortYN = true;
                        }
                    }

                    if (itemBSortYN == true) //BaseForm에 ItmBSort(품목대분류) 라는 컨트롤이 존재할 경우만 코드 연동
                    {
                        oForm.Items.Item("ItmBsort").Specific.Value = oBaseForm01.Items.Item("ItmBSort").Specific.Value; //BaseForm의 제품대분류 코드 연동
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
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
        /// PS_SM010_MTX01
        /// </summary>
        private void PS_SM010_MTX01()
        {
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
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("StockType").Specific.Selected.Value.ToString().Trim();
                Param03 = oForm.Items.Item("TradeType").Specific.Selected.Value.ToString().Trim();
                Param04 = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim();
                Param05 = oForm.Items.Item("ItmMsort").Specific.Value.ToString().Trim();
                Param06 = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
                Param07 = oForm.Items.Item("ItemType").Specific.Selected.Value.ToString().Trim();
                Param08 = oForm.Items.Item("Mark").Specific.Selected.Value.ToString().Trim();
                Param09 = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
                Param10 = oForm.Items.Item("ItemGpCd").Specific.Selected.Value.ToString().Trim();
                Param11 = oForm.Items.Item("WhsCode").Specific.Selected.Value.ToString().Trim();

                if (oBaseForm01 == null)
                {
                    Query01 = "EXEC PS_SM010_01 '" + Param01 + "','','','" + Param02 + "','','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "','" + Param11 + "'";
                }
                else if (oBaseForm01.Type == Convert.ToDouble("149") || oBaseForm01.Type == Convert.ToDouble("139") || oBaseForm01.Type == Convert.ToDouble("140") || oBaseForm01.Type == Convert.ToDouble("180") || oBaseForm01.Type == Convert.ToDouble("133") || oBaseForm01.Type == Convert.ToDouble("179") || oBaseForm01.Type == Convert.ToDouble("60091"))
                {
                    Query01 = "EXEC PS_SM010_01 '" + Param01 + "','Y','','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "','" + Param11 + "'";
                    //판매Y,구매,재고타입(1:재고있는것만,2:전체),거래타입(1:일반,2:임가공)
                }
                else if (oBaseForm01.Type == Convert.ToDouble("142") || oBaseForm01.Type == Convert.ToDouble("143") || oBaseForm01.Type == Convert.ToDouble("182") || oBaseForm01.Type == Convert.ToDouble("141") || oBaseForm01.Type == Convert.ToDouble("181") || oBaseForm01.Type == Convert.ToDouble("60092"))
                {
                    Query01 = "EXEC PS_SM010_01 '" + Param01 + "','','Y','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "','" + Param11 + "'";
                    //판매,구매Y
                }
                else
                {
                    Query01 = "EXEC PS_SM010_01 '" + Param01 + "','','','" + Param02 + "','','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "','" + Param11 + "'";
                }
                oGrid01.DataTable.Clear();

                oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(Query01);
                oGrid01.DataTable = oForm.DataSources.DataTables.Item("DataTable");

                if (oGrid01.Rows.Count == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                oGrid01.AutoResizeColumns();
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
        private void PS_SM010_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SM010'", "");
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_SM010_SetBaseForm
        /// </summary>
        private void PS_SM010_SetBaseForm()
        {
            int i = 0;
            int Today_Renamed;
            string sQry;
            string ItemCode;
            string errMessage = string.Empty;
            SAPbouiCOM.Matrix oBaseMat01;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (oBaseForm01 == null)
                {
                }

                // csee 1 별도
                else if (oBaseForm01.TypeEx == "721")
                {
                    oBaseMat01 = oBaseForm01.Items.Item("13").Specific;
                    for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                    {
                        oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                        oBaseColRow01 += 1;
                    }
                }

                // csee 2 (BItemCode - BitemCod / Mat01 - ItemCode) 별도
                else if (oBaseForm01.TypeEx == "PS_MM180")
                {
                    if (oBaseItemUID01 == "BItemCod")
                    {
                        oBaseForm01.Items.Item("BItemCod").Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                    }
                    else if (oBaseItemUID01 == "Mat01")
                    {
                        oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
                        for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                        {
                            oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            oBaseColRow01 += 1;
                        }
                    }
                }

                //case 3 별도
                else if (oBaseForm01.TypeEx == "PS_MM007")
                {
                    sQry = "  SELECT    CASE";
                    sQry += " WHEN LEFT(U_ItmBsort, 1) = '4' THEN '20'";
                    sQry += " WHEN LEFT(U_ItmBsort, 1) = '3' THEN '10'";
                    sQry += " WHEN LEFT(U_ItmBsort, 1) = '2' THEN '50'";
                    sQry += " END";
                    sQry += " FROM      OITM";
                    sQry += " WHERE     ItemCode = '" + oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value + "'";

                    oRecordSet01.DoQuery(sQry);
                    
                    Today_Renamed = Convert.ToInt32(DateTime.Now.ToString("yyyyMMdd"));

                    if (Today_Renamed >= 20150701)
                    {
                        //선택된행의수
                        for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                        {
                            ItemCode = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            if (oRecordSet01.Fields.Item(0).Value.ToString().Trim() == "20")
                            {
                                if (ItemCode.Length != 11)
                                {
                                    oForm.DataSources.UserDataSources.Item("Check01").Value = "Y";
                                    errMessage = "통합부자재 코드로만 선택이 가능합니다. 확인바랍니다.";
                                    throw new Exception();
                                }
                            }
                        }
                    }

                    if (oBaseItemUID01 == "Mat01")
                    {
                        oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
                        for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                        {
                            oBaseMat01.Columns.Item("MATNR").Cells.Item(oBaseColRow01).Specific.Value = codeHelpClass.Left(oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value, 7);
                            oBaseMat01.Columns.Item("MAKTX").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품명").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value + " : ";
                            oBaseMat01.Columns.Item("MEINS").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("단위").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            oBaseMat01.Columns.Item("MATKL").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("분류코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            oBaseColRow01 += 1;
                        }
                    }
                }

                // csee 4 별도
                else if (oBaseForm01.TypeEx == "PS_MM005")
                {
                    sQry = "  SELECT    CASE";
                    sQry += "               WHEN LEFT(U_ItmBsort, 1) = '4' THEN '20'";
                    sQry += "               WHEN LEFT(U_ItmBsort, 1) = '3' THEN '10'";
                    sQry += "               WHEN LEFT(U_ItmBsort, 1) = '2' THEN '50'";
                    sQry += "           END";
                    sQry += " FROM      OITM";
                    sQry += " WHERE     ItemCode = '" + oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value + "'";

                    oRecordSet01.DoQuery(sQry);

                    sQry02 = " Select convert(char(8),GetDate(),112) ";
                    oRecordset02.DoQuery(sQry02);
                    Today_Renamed = oRecordset02.Fields.Item(0).Value.ToString().Trim();

                    if (oBaseForm01.Items.Item("OrdType").Specific.Value == "10" || oBaseForm01.Items.Item("OrdType").Specific.Value == "20" || oBaseForm01.Items.Item("OrdType").Specific.Value == "50")
                    {
                        if (oRecordSet01.Fields.Item(0).Value == oBaseForm01.Items.Item("OrdType").Specific.Value)
                        {
                            if (oBaseItemUID01 == "ItemCode")
                            {
                                oBaseForm01.Items.Item("ItemCode").Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            }
                            else if (oBaseItemUID01 == "Mat01")
                            {
                                oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
                                for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                                {
                                    oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                                    oBaseColRow01 += 1;
                                }
                            }
                        }
                        else
                        {
                            errMessage = "품의와 품목구분이 일치하지않아 등록할 수 없습니다. 확인하십시오.";
                            throw new Exception();
                        }
                    }
                    else
                    {
                        if (oBaseItemUID01 == "ItemCode")
                        {
                            oBaseForm01.Items.Item("ItemCode").Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                        }
                        else if (oBaseItemUID01 == "Mat01")
                        {
                            oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
                            for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                            {
                                oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                                oBaseColRow01 += 1;
                            }
                        }
                    }
                }

                //case 5 Mat01 - OrdMgNum
                else if (oBaseForm01.TypeEx == "PS_PP043")
                {
                    if (oBaseItemUID01 == "ItemCode")
                    {
                    }
                    else if (oBaseItemUID01 == "Mat01")
                    {
                        oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
                        for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                        {
                            oBaseMat01.Columns.Item("OrdMgNum").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            oBaseColRow01 += 1;
                        }
                    }
                }

                //case 6 ItemCode - ItemCode / Mat01 - MItemCod
                else if (oBaseForm01.TypeEx == "PS_MM002") //별도
                {
                    if (oBaseItemUID01 == "ItemCode")
                    {
                        oBaseForm01.Items.Item("ItemCode").Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                    }
                    else if (oBaseItemUID01 == "Mat01")
                    {
                        oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
                        for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                        {
                            oBaseMat01.Columns.Item("MItemCod").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            oBaseColRow01 += 1;
                        }
                    }
                }

                //case 7 Mat01 ItmMsort
                else if (oBaseForm01.TypeEx == "PS_PP130" || oBaseForm01.TypeEx == "PS_PP940")
                {
                    if (oBaseItemUID01 == "Mat01")
                    {
                        oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
                        for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                        {
                            if (string.IsNullOrEmpty(oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value))
                            {
                                oBaseMat01.Columns.Item("ItmMsort").Cells.Item(oBaseColRow01).Specific.Value = codeHelpClass.Mid(oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value, 0, 5);
                            }
                            oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            oBaseColRow01 += 1;
                        }
                    }
                }

                //case 8 Mat01 ItmBsort,ItmMsort
                else if (oBaseForm01.TypeEx == "PS_QM903") // 별도
                {
                    if (oBaseItemUID01 == "Mat01")
                    {
                        oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
                        for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                        {
                            if (string.IsNullOrEmpty(oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value))
                            {
                                oBaseMat01.Columns.Item("ItmBsort").Cells.Item(oBaseColRow01).Specific.Value = codeHelpClass.Mid(oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value, 1, 3);
                                oBaseMat01.Columns.Item("ItmMsort").Cells.Item(oBaseColRow01).Specific.Value = codeHelpClass.Mid(oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value, 1, 5);
                            }
                            oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            oBaseColRow01 += 1;
                        }
                    }
                }
                
                //case 9 (ItemCode - ItemCode, Mat01 - ItemCode) 별도
                else if (oBaseForm01.TypeEx == "PS_PP950" || oBaseForm01.TypeEx == "PS_QM010" || oBaseForm01.TypeEx == "PS_PP077" || oBaseForm01.TypeEx == "PS_PP078" || oBaseForm01.TypeEx == "PS_QM005" || oBaseForm01.TypeEx == "PS_SD091" || 
                         oBaseForm01.TypeEx == "PS_QM011" || oBaseForm01.TypeEx == "PS_MM100" || oBaseForm01.TypeEx == "PS_PP025" || oBaseForm01.TypeEx == "PS_SD006" || oBaseForm01.TypeEx == "PS_SD005" || oBaseForm01.TypeEx == "PS_PP083" || 
                         oBaseForm01.TypeEx == "PS_MM097" || oBaseForm01.TypeEx == "PS_QM060" || oBaseForm01.TypeEx == "PS_PP020" || oBaseForm01.TypeEx == "PS_PP010" || oBaseForm01.TypeEx == "PS_MM030" || oBaseForm01.TypeEx == "PS_MM070" || 
                         oBaseForm01.TypeEx == "PS_SD010" || oBaseForm01.TypeEx == "PS_MM110")
                {
                    if (oBaseItemUID01 == "ItemCode")
                    {
                        oBaseForm01.Items.Item("ItemCode").Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                    }
                    else if (oBaseItemUID01 == "Mat01")
                    {
                        oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
                        for (i = 0; i <= oGrid01.Rows.SelectedRows.Count - 1; i++)
                        {
                            oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Value;
                            oBaseColRow01 += 1;
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_SM010_DataValidCheck()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;
            try
            {
                PS_SM010_FormClear();
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

            }
            return functionReturnValue;
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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_SM010_MTX01();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "Button02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_SM010_SetBaseForm();
                            if (oForm.DataSources.UserDataSources.Item("Check01").Value.ToString().Trim() == "N")
                            {
                                oForm.Close();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemName", "");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItmBsort", "");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItmMsort", "");
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02")
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
        }

        /// <summary>
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string oQuery01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "ItmBsort")
                        {
                            oQuery01 = "SELECT Name FROM [@PSH_ITMBSORT] WHERE Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            oRecordSet01.DoQuery(oQuery01);
                            oForm.Items.Item("ItmBName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (pVal.ItemUID == "ItmMsort")
                        {
                            oQuery01 = "SELECT U_CodeName FROM [@PSH_ITMMSORT] WHERE U_Code = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            oRecordSet01.DoQuery(oQuery01);
                            oForm.Items.Item("ItmMName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
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
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.Row > 0)
                            {
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.Row >= 0)
                            {
                                oForm.Items.Item("LItemCod").Specific.Value = oGrid01.DataTable.Columns.Item("품목코드").Cells.Item(oGrid01.GetDataTableRowIndex(pVal.Row)).Value.ToString().Trim();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_SM010_FormItemEnabled();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (oDataTable01 != null) //SelectedObjects 가 null이 아닐때만 실행(ChooseFromList 팝업창을 취소했을 때 미실행)
                    {
                        if (pVal.ItemUID == "ItemCode")
                        {
                            oForm.DataSources.UserDataSources.Item("ItemCode").Value = oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value;
                            oForm.DataSources.UserDataSources.Item("ItemName").Value = oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value;
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
                oForm.Update();
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (pVal.Row == -1)
                        {
                        }
                        else
                        {
                            if (oGrid01.Rows.SelectedRows.Count > 0)
                            {
                                PS_SM010_SetBaseForm();
                                //부모폼에입력
                                if (oForm.DataSources.UserDataSources.Item("Check01").Value.ToString().Trim() == "N")
                                {
                                    oForm.Close();
                                }
                            }
                            else
                            {
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
                if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02")
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
    }
}
