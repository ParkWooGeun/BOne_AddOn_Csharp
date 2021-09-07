using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 작번등록
    /// </summary>
    internal class PS_PP020 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.DBDataSource oDS_PS_PP020H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_TEMPTABLE;
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private string oLastRightClickDocEntry;

        /// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP020.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP020_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP020");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                
                PS_PP020_CreateItems();
                PS_PP020_SetComboBox();
                PS_PP020_AddMatrixRow(0, true, "");
                PS_PP020_LoadCaption();
                PS_PP020_EnableFormItem();

                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1287", false); //복제
                oForm.EnableMenu("1285", false); //복원
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", true); //행삭제
                oForm.EnableMenu("1299", true); //행닫기
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_PP020_CreateItems()
        {
            try
            {
                oDS_PS_PP020H = oForm.DataSources.DBDataSources.Item("@PS_PP020H");
                oDS_PS_TEMPTABLE = oForm.DataSources.DBDataSources.Item("@PS_TEMPTABLE");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.AutoResizeColumns();

                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat02.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("PuDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("PuDateFr").Specific.DataBind.SetBound(true, "", "PuDateFr");

                oForm.DataSources.UserDataSources.Add("PuDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("PuDateTo").Specific.DataBind.SetBound(true, "", "PuDateTo");

                oForm.DataSources.UserDataSources.Add("JakDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("JakDateFr").Specific.DataBind.SetBound(true, "", "JakDateFr");

                oForm.DataSources.UserDataSources.Add("JakDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("JakDateTo").Specific.DataBind.SetBound(true, "", "JakDateTo");

                oForm.DataSources.UserDataSources.Add("RadioMat01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("RadioMat01").Specific.DataBind.SetBound(true, "", "RadioMat01");

                oForm.DataSources.UserDataSources.Add("RadioMat02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("RadioMat02").Specific.DataBind.SetBound(true, "", "RadioMat02");

                oForm.Items.Item("RadioMat01").Specific.GroupWith("RadioMat02");

                oMat01.Columns.Item("OrderAmt").Visible = false; //메인작번등록-견적금액 숨김
                oMat01.Columns.Item("NegoAmt").Visible = false; //메인작번등록-네고금액 숨김
                oMat01.Columns.Item("TrgtAmt").Visible = false; //메인작번등록-목표금액 숨김

                oMat02.Columns.Item("OrderAmt").Visible = false; //서브작번등록-견적금액 숨김
                oMat02.Columns.Item("NegoAmt").Visible = false; //서브작번등록-네고금액 숨김
                oMat02.Columns.Item("TrgtAmt").Visible = false; //서브작번등록-목표금액 숨김
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP020_SetComboBox()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                //품목대분류
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where Code in ('105', '106') Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("ItmBSort").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("ItmBSort").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("ItmBSort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //작업구분
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("10", "영업");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("20", "정비");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("30", "멀티");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("40", "소재");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("50", "R/D");
                oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("60", "견본");
                oForm.Items.Item("WorkGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oMat01.Columns.Item("WorkGbn").ValidValues.Add("10", "영업");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("20", "정비");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("30", "멀티");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("40", "소재");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("50", "R/D");
                oMat01.Columns.Item("WorkGbn").ValidValues.Add("60", "견본");

                oMat02.Columns.Item("WorkGbn").ValidValues.Add("10", "영업");
                oMat02.Columns.Item("WorkGbn").ValidValues.Add("20", "정비");
                oMat02.Columns.Item("WorkGbn").ValidValues.Add("30", "멀티");
                oMat02.Columns.Item("WorkGbn").ValidValues.Add("40", "소재");
                oMat02.Columns.Item("WorkGbn").ValidValues.Add("50", "R/D");
                oMat02.Columns.Item("WorkGbn").ValidValues.Add("60", "견본");

                //작번구분
                oForm.Items.Item("JakGbn").Specific.ValidValues.Add("99", "전체");
                oForm.Items.Item("JakGbn").Specific.ValidValues.Add("00", "메인작번");
                oForm.Items.Item("JakGbn").Specific.ValidValues.Add("01", "서브작번");
                oForm.Items.Item("JakGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //재작업구분
                oForm.Items.Item("ReWork").Specific.ValidValues.Add("10", "정상");
                oForm.Items.Item("ReWork").Specific.ValidValues.Add("20", "재작업");
                oForm.Items.Item("ReWork").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //작번구분
                oMat01.Columns.Item("InOutGbn").ValidValues.Add("%", "선택");
                oMat01.Columns.Item("InOutGbn").ValidValues.Add("IN", "자체");
                oMat01.Columns.Item("InOutGbn").ValidValues.Add("OUT", "외주");

                oMat02.Columns.Item("InOutGbn").ValidValues.Add("%", "선택");
                oMat02.Columns.Item("InOutGbn").ValidValues.Add("IN", "자체");
                oMat02.Columns.Item("InOutGbn").ValidValues.Add("OUT", "외주");

                //재작업 사유(Mat01)
                sQry = "  SELECT  B.U_Minor, ";
                sQry += "         B.U_CdName";
                sQry += " FROM    [@PS_SY001H] AS A";
                sQry += "         INNER JOIN";
                sQry += "         [@PS_SY001L] AS B";
                sQry += "             ON A.Code = B.Code";
                sQry += "             AND A.Code = 'P202'";

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ReWkRn"), sQry, "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat02.Columns.Item("ReWkRn"), sQry, "", ""); //재작업 사유(Mat02)

                //연간품여부
                oMat01.Columns.Item("YearPdYN").ValidValues.Add("N", "N");
                oMat01.Columns.Item("YearPdYN").ValidValues.Add("Y", "Y");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 행 추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        /// <param name="ItemUID"></param>
        private void PS_PP020_AddMatrixRow(int oRow, bool RowIserted, string ItemUID)
        {
            try
            {
                if (ItemUID == "Mat02")
                {
                    if (RowIserted == false)
                    {
                        oDS_PS_TEMPTABLE.InsertRecord(oRow);
                    }
                    oMat02.AddRow();
                    oDS_PS_TEMPTABLE.Offset = oRow;
                    oDS_PS_TEMPTABLE.SetValue("U_iField01", oRow, Convert.ToString(oRow + 1));
                    oMat02.LoadFromDataSource();
                }
                else
                {
                    if (RowIserted == false)
                    {
                        oDS_PS_PP020H.InsertRecord(oRow);
                    }
                    oMat01.AddRow();
                    oDS_PS_PP020H.Offset = oRow;
                    oDS_PS_PP020H.SetValue("DocNum", oRow, Convert.ToString(oRow + 1));
                    oMat01.LoadFromDataSource();
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 버튼 캡션 설정
        /// </summary>
        private void PS_PP020_LoadCaption()
        {
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                oForm.Items.Item("Btn01").Specific.Caption = "추가";
            }
            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                oForm.Items.Item("Btn01").Specific.Caption = "확인";
            }
            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                oForm.Items.Item("Btn01").Specific.Caption = "갱신";
            }
        }

        /// <summary>
        /// 각모드에따른 아이템설정
        /// </summary>
        private void PS_PP020_EnableFormItem()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("ItmBSort").Enabled = true;
                    oForm.Items.Item("WorkGbn").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("PuDateFr").Enabled = true;
                    oForm.Items.Item("PuDateTo").Enabled = true;
                    oForm.Items.Item("JakDateFr").Enabled = false;
                    oForm.Items.Item("JakDateTo").Enabled = false;
                    oForm.Items.Item("JakGbn").Enabled = false;
                    oForm.Items.Item("Btn02").Enabled = false;

                    oMat01.Columns.Item("JakName").Editable = true;
                    oMat01.Columns.Item("SubNo1").Editable = false;
                    oMat01.Columns.Item("SubNo2").Editable = false;
                    oMat01.Columns.Item("JakMyung").Editable = true;
                    oMat01.Columns.Item("JakSize").Editable = true;
                    oMat01.Columns.Item("JakUnit").Editable = true;
                    oMat01.Columns.Item("RegNum").Editable = false;
                    oMat01.Columns.Item("ItemCode").Editable = false;
                    oMat01.Columns.Item("CardCode").Editable = true;
                    oMat01.Columns.Item("ShipCode").Editable = true;
                    oMat01.Columns.Item("InOutGbn").Editable = true;
                    oMat01.Columns.Item("ProDate").Editable = true;
                    oMat01.Columns.Item("ReDate").Editable = true;
                    oMat01.Columns.Item("WrWeight").Editable = true;
                    oMat01.Columns.Item("Comments").Editable = true;
                    oMat01.Columns.Item("Status").Editable = true;

                    oForm.Items.Item("RadioMat02").Visible = false;
                    oForm.Items.Item("Mat02").Visible = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("ItmBSort").Enabled = true;
                    oForm.Items.Item("WorkGbn").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("PuDateFr").Enabled = false;
                    oForm.Items.Item("PuDateTo").Enabled = false;
                    oForm.Items.Item("JakDateFr").Enabled = true;
                    oForm.Items.Item("JakDateTo").Enabled = true;
                    oForm.Items.Item("JakGbn").Enabled = true;
                    oForm.Items.Item("Btn02").Enabled = true;

                    oMat01.Columns.Item("JakName").Editable = false;
                    oMat01.Columns.Item("SubNo1").Editable = false;
                    oMat01.Columns.Item("SubNo2").Editable = false;
                    oMat01.Columns.Item("JakMyung").Editable = true;
                    oMat01.Columns.Item("JakSize").Editable = true;
                    oMat01.Columns.Item("JakUnit").Editable = true;
                    oMat01.Columns.Item("RegNum").Editable = false;
                    oMat01.Columns.Item("ItemCode").Editable = false;
                    oMat01.Columns.Item("CardCode").Editable = true;
                    oMat01.Columns.Item("ShipCode").Editable = true;
                    oMat01.Columns.Item("InOutGbn").Editable = true;
                    oMat01.Columns.Item("ProDate").Editable = true;
                    oMat01.Columns.Item("ReDate").Editable = true;
                    oMat01.Columns.Item("WrWeight").Editable = true;
                    oMat01.Columns.Item("Comments").Editable = true;
                    oMat01.Columns.Item("Status").Editable = true;

                    oForm.Items.Item("RadioMat02").Visible = true;
                    oForm.Items.Item("Mat02").Visible = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FlushToItemValue
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_PP020_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            string sQry;
            string JakName;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "ShipCode":
                        sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("ShipCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ShipName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "ItemCode":
                        sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ItemName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "CardCode":
                        sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("CardName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "Mat01":
                        if (oCol == "JakName")
                        {
                            oForm.Freeze(true);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("JakName").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_PP020_AddMatrixRow(oMat01.RowCount, false, "");
                                }
                            }

                            JakName = oMat01.Columns.Item("JakName").Cells.Item(oRow).Specific.Value.ToString().Trim();
                            sQry = "   Select  a.DocEntry,";
                            sQry +=  "         a.DocNum,";
                            sQry +=  "         a.Period,";
                            sQry +=  "         a.Instance,";
                            sQry +=  "         a.Series,";
                            sQry +=  "         a.Handwrtten,";
                            sQry +=  "         a.Canceled,";
                            sQry +=  "         a.Object,";
                            sQry +=  "         a.LogInst,";
                            sQry +=  "         a.UserSign,";
                            sQry +=  "         a.Transfered,";
                            sQry +=  "         a.Status,";
                            sQry +=  "         a.CreateDate,";
                            sQry +=  "         a.CreateTime,";
                            sQry +=  "         a.UpdateDate,";
                            sQry +=  "         a.UpdateTime,";
                            sQry +=  "         a.DataSource,";
                            sQry +=  "         a.U_BPLId,";
                            sQry +=  "         a.U_RegNum,";
                            sQry +=  "         a.U_ItemCode,";
                            sQry +=  "         a.U_ItemName,";
                            sQry +=  "         a.U_Material,";
                            sQry +=  "         a.U_Unit,";
                            sQry +=  "         a.U_Size,";
                            sQry +=  "         a.U_ItmBSort,";
                            sQry +=  "         a.U_CpName,";
                            sQry +=  "         a.U_SjDocNum,";
                            sQry +=  "         a.U_SjLinNum,";
                            sQry +=  "         a.U_SjQty,";
                            sQry +=  "         a.U_SjWeight,";
                            sQry +=  "         a.U_SjDcDate,";
                            sQry +=  "         a.U_SjDuDate,";
                            sQry +=  "         a.U_SlePrice,";
                            sQry +=  "         a.U_WorkGbn,";
                            sQry +=  "         a.U_CardCode,";
                            sQry +=  "         a.U_CardName,";
                            sQry +=  "         a.U_ShipCode,";
                            sQry +=  "         a.U_ShipName,";
                            sQry +=  "         a.U_PuDate,";
                            sQry +=  "         a.U_ProDate,";
                            sQry +=  "         a.U_Comments,";
                            sQry +=  "         a.U_JakName,";
                            sQry +=  "         a.U_SubNo1,";
                            sQry +=  "         a.U_SubNo2,";
                            sQry +=  "         a.U_Status,";
                            sQry +=  "         U_ReqCod = '" + dataHelpClass.User_MSTCOD() + "',";
                            sQry +=  "         a.U_UseDept,";
                            sQry +=  "         ISNULL(a.U_ReWeight, 0) AS U_ReWeight";
                            sQry +=  " FROM    [@PS_PP010H] a ";
                            sQry +=  "         LEFT JOIN ";
                            sQry +=  "         [@PS_SD010H] b ";
                            sQry +=  "             ON a.U_RegNum = b.U_RegNum";
                            sQry +=  " WHERE   a.U_JakName = '" + JakName + "'";

                            oRecordSet01.DoQuery(sQry);

                            if (!string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value.ToString().Trim()))
                            {
                                sQry = "   Select  a.DocEntry,";
                                sQry +=  "         a.DocNum,";
                                sQry +=  "         a.Period,";
                                sQry +=  "         a.Instance,";
                                sQry +=  "         a.Series,";
                                sQry +=  "         a.Handwrtten,";
                                sQry +=  "         a.Canceled,";
                                sQry +=  "         a.Object,";
                                sQry +=  "         a.LogInst,";
                                sQry +=  "         a.UserSign,";
                                sQry +=  "         a.Transfered,";
                                sQry +=  "         a.Status,";
                                sQry +=  "         a.CreateDate,";
                                sQry +=  "         a.CreateTime,";
                                sQry +=  "         a.UpdateDate,";
                                sQry +=  "         a.UpdateTime,";
                                sQry +=  "         a.DataSource,";
                                sQry +=  "         a.U_BPLId,";
                                sQry +=  "         a.U_RegNum,";
                                sQry +=  "         a.U_ItemCode,";
                                sQry +=  "         a.U_ItemName,";
                                sQry +=  "         a.U_Material,";
                                sQry +=  "         a.U_Unit,";
                                sQry +=  "         a.U_Size,";
                                sQry +=  "         a.U_ItmBSort,";
                                sQry +=  "         a.U_CpName,";
                                sQry +=  "         a.U_SjDocNum,";
                                sQry +=  "         a.U_SjLinNum,";
                                sQry +=  "         a.U_SjQty,";
                                sQry +=  "         a.U_SjWeight,";
                                sQry +=  "         a.U_SjDcDate,";
                                sQry +=  "         a.U_SjDuDate,";
                                sQry +=  "         a.U_SlePrice,";
                                sQry +=  "         a.U_WorkGbn,";
                                sQry +=  "         a.U_CardCode,";
                                sQry +=  "         a.U_CardName,";
                                sQry +=  "         a.U_ShipCode,";
                                sQry +=  "         a.U_ShipName,";
                                sQry +=  "         a.U_PuDate,";
                                sQry +=  "         a.U_ProDate,";
                                sQry +=  "         a.U_Comments,";
                                sQry +=  "         a.U_JakName,";
                                sQry +=  "         a.U_SubNo1,";
                                sQry +=  "         a.U_SubNo2,";
                                sQry +=  "         a.U_Status,";
                                sQry +=  "         U_ReqCod = '" + dataHelpClass.User_MSTCOD() + "',";
                                sQry +=  "         a.U_UseDept,";
                                sQry +=  "         ISNULL(a.U_ReWeight, 0) AS U_ReWeight,";
                                sQry +=  "         B.SalUnitMsr ";
                                sQry +=  " FROM    [@PS_PP010H] a ";
                                sQry +=  "         INNER JOIN ";
                                sQry +=  "         [OITM] b ";
                                sQry +=  "             ON a.U_ItemCode = b.ItemCode ";
                                sQry +=  " WHERE   a.U_jakName = '" + JakName + "'";

                                oRecordSet01.DoQuery(sQry);
                            }

                            oMat01.Columns.Item("SubNo1").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SubNo1").Value.ToString().Trim();
                            oMat01.Columns.Item("SubNo2").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SubNo2").Value.ToString().Trim();
                            oMat01.Columns.Item("RegNum").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_RegNum").Value.ToString().Trim();
                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim();
                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim();
                            oMat01.Columns.Item("Material").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_Material").Value.ToString().Trim();
                            oMat01.Columns.Item("Unit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("SalUnitMsr").Value.ToString().Trim();
                            oMat01.Columns.Item("Size").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_Size").Value.ToString().Trim();
                            oMat01.Columns.Item("ItmBSort").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("U_ItmBSort").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat01.Columns.Item("SjDocNum").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SjDocNum").Value.ToString().Trim();
                            oMat01.Columns.Item("SjLinNum").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SjLinNum").Value.ToString().Trim();
                            oMat01.Columns.Item("CardCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim();
                            oMat01.Columns.Item("CardName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim();
                            oMat01.Columns.Item("ShipCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ShipCode").Value.ToString().Trim();
                            oMat01.Columns.Item("ShipName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ShipName").Value.ToString().Trim();
                            oMat01.Columns.Item("InOutGbn").Cells.Item(oRow).Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat01.Columns.Item("JakDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat01.Columns.Item("ProDate").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ProDate").Value.ToString("yyyyMMdd");
                            oMat01.Columns.Item("ReDate").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SjDuDate").Value.ToString("yyyyMMdd");
                            oMat01.Columns.Item("SjWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SjWeight").Value.ToString().Trim();
                            oMat01.Columns.Item("SjDcDate").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SjDcDate").Value.ToString("yyyyMMdd");
                            oMat01.Columns.Item("SjDuDate").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SjDuDate").Value.ToString("yyyyMMdd");
                            oMat01.Columns.Item("SlePrice").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_SlePrice").Value.ToString().Trim();
                            oMat01.Columns.Item("WorkGbn").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("U_WorkGbn").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat01.Columns.Item("PP010Doc").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim();
                            oMat01.Columns.Item("ReqCod").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ReqCod").Value.ToString().Trim();
                            oMat01.Columns.Item("WrWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ReWeight").Value.ToString().Trim();
                            oMat01.Columns.Item("YearPdYN").Cells.Item(oRow).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);

                            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + dataHelpClass.User_MSTCOD() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ReqNam").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                            oMat01.Columns.Item("Comments").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Freeze(false);
                        }
                        else if (oCol == "ItemCode")
                        {
                            oForm.Freeze(true);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_PP020_AddMatrixRow(oMat01.RowCount, false, "");
                                }
                            }

                            sQry = "   Select  a.ItemName, ";
                            sQry +=  "         a.U_Material, ";
                            sQry +=  "         a.SalUnitMsr, ";
                            sQry +=  "         a.U_Size, ";
                            sQry +=  "         a.U_ItmBSort, ";
                            sQry +=  "         b.U_CPNaming  ";
                            sQry +=  " From    OITM a ";
                            sQry +=  "         Left Join ";
                            sQry +=  "         [@PSH_ItmMSort] b";
                            sQry +=  "             ON a.U_ItmBSort = b.U_rCode ";
                            sQry +=  "             And a.U_ItmMsort = b.U_Code ";
                            sQry +=  " WHERE   a.ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Freeze(false);
                        }
                        else if (oCol == "ShipCode")
                        {
                            sQry = "Select CardName From OCRD Where CardCode = '" + oMat01.Columns.Item("ShipCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ShipName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (oCol == "ReqCod")
                        {
                            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oMat01.Columns.Item("ReqCod").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ReqNam").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        oMat01.Columns.Item(oCol).Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oMat01.AutoResizeColumns();
                        
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PS_PP020_LoadData()
        {
            short i;
            string sQry;
            string ItmBsort;
            string ItemCode;
            string CardCode;
            string WorkGbn;
            string JakDateTo;
            string BPLId;
            string JakDateFr;
            string JakGbn;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
                WorkGbn = oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim();
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                JakDateFr = oForm.Items.Item("JakDateFr").Specific.Value.ToString().Trim();
                JakDateTo = oForm.Items.Item("JakDateTo").Specific.Value.ToString().Trim();
                JakGbn = oForm.Items.Item("JakGbn").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(BPLId))
                {
                    BPLId = "%";
                }
                    
                if (string.IsNullOrEmpty(CardCode))
                {
                    CardCode = "%";
                }
                    
                if (string.IsNullOrEmpty(ItemCode))
                {
                    ItemCode = "%";
                }
                    
                if (string.IsNullOrEmpty(JakDateFr))
                {
                    JakDateFr = "19000101";
                }
                    
                if (string.IsNullOrEmpty(JakDateTo))
                {
                    JakDateTo = "20991231";
                }
                    
                if (string.IsNullOrEmpty(JakGbn))
                {
                    JakGbn = "%";
                }

                sQry = "EXEC [PS_PP020_01] '";
                sQry += BPLId + "','";
                sQry += ItmBsort + "','";
                sQry += WorkGbn + "','";
                sQry += CardCode + "','";
                sQry += ItemCode + "','";
                sQry += JakDateFr + "','";
                sQry += JakDateTo + "','";
                sQry += JakGbn + "'";

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_PP020H.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_PP020H.Size)
                    {
                        oDS_PS_PP020H.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_PP020H.Offset = i;
                    oDS_PS_PP020H.SetValue("DocNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP020H.SetValue("DocEntry", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_JakName", i, oRecordSet01.Fields.Item("U_JakName").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_SubNo1", i, oRecordSet01.Fields.Item("U_SubNo1").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_SubNo2", i, oRecordSet01.Fields.Item("U_SubNo2").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_JakMyung", i, oRecordSet01.Fields.Item("U_JakMyung").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_JakSize", i, oRecordSet01.Fields.Item("U_JakSize").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_JakUnit", i, oRecordSet01.Fields.Item("U_JakUnit").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_RegNum", i, oRecordSet01.Fields.Item("U_RegNum").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_Material", i, oRecordSet01.Fields.Item("U_Material").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_Unit", i, oRecordSet01.Fields.Item("U_Unit").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_Size", i, oRecordSet01.Fields.Item("U_Size").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_ItmBSort", i, oRecordSet01.Fields.Item("U_ItmBSort").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_SjDocNum", i, oRecordSet01.Fields.Item("U_SjDocNum").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_SjLinNum", i, oRecordSet01.Fields.Item("U_SjLinNum").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_SjWeight", i, oRecordSet01.Fields.Item("U_SjWeight").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_WrWeight", i, oRecordSet01.Fields.Item("U_WrWeight").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_SjDcDate", i, oRecordSet01.Fields.Item("U_SjDcDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP020H.SetValue("U_SjDuDate", i, oRecordSet01.Fields.Item("U_SjDuDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP020H.SetValue("U_SlePrice", i, oRecordSet01.Fields.Item("U_SlePrice").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_WorkGbn", i, oRecordSet01.Fields.Item("U_WorkGbn").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_CardCode", i, oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_CardName", i, oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_ShipCode", i, oRecordSet01.Fields.Item("U_ShipCode").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_ShipName", i, oRecordSet01.Fields.Item("U_ShipName").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_JakDate", i, oRecordSet01.Fields.Item("U_JakDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP020H.SetValue("U_ProDate", i, oRecordSet01.Fields.Item("U_ProDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP020H.SetValue("U_ReDate", i, oRecordSet01.Fields.Item("U_ReDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_PP020H.SetValue("U_Comments", i, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_InOutGbn", i, oRecordSet01.Fields.Item("U_InOutGbn").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_Status", i, oRecordSet01.Fields.Item("U_Status").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_ReqCod", i, oRecordSet01.Fields.Item("U_ReqCod").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_ReqNam", i, oRecordSet01.Fields.Item("U_ReqNam").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_PP010Doc", i, oRecordSet01.Fields.Item("U_PP010Doc").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_DrawQty", i, oRecordSet01.Fields.Item("U_DrawQty").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_ReWkRn", i, oRecordSet01.Fields.Item("U_ReWkRn").Value.ToString().Trim());
                    oDS_PS_PP020H.SetValue("U_YearPdYN", i, oRecordSet01.Fields.Item("U_YearPdYN").Value.ToString().Trim());

                    oDS_PS_PP020H.SetValue("U_OrderAmt", i, oRecordSet01.Fields.Item("U_OrderAmt").Value.ToString().Trim()); //수주금액
                    oDS_PS_PP020H.SetValue("U_NegoAmt", i, oRecordSet01.Fields.Item("U_NegoAmt").Value.ToString().Trim()); //Nego금액
                    oDS_PS_PP020H.SetValue("U_TrgtAmt", i, oRecordSet01.Fields.Item("U_TrgtAmt").Value.ToString().Trim()); //목표금액

                    oRecordSet01.MoveNext();
                    ProgBar01.Value += 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
            }
        }

        /// <summary>
        /// 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
        /// </summary>
        /// <returns></returns>
        private bool PS_PP020_AddData()
        {
            bool returnValue = false;
            short loopCount;
            string errMessage = string.Empty;
            string ErrOrdNum = null; //선행프로세스보다 일자가 빨라서 저장되지 않는 작번을 저장
            string Query01;
            string Query02;
            string DocEntry;
            string BPLId;
            string JakName;
            string SubNo1;
            string SubNo2;
            string JakMyung;
            string JakSize;
            string JakUnit;
            string RegNum;
            string ItemCode;
            string ItemName;
            string Material;
            string Unit;
            string Size;
            string ItmBsort;
            string SjDocNum;
            string SjLinNum;
            string SjWeight;
            string WrWeight;
            string SjDcDate;
            string SjDuDate;
            string SlePrice;
            string WorkGbn;
            string CardCode;
            string CardName;
            string ShipCode;
            string ShipName;
            string InOutGbn;
            string JakDate;
            string ProDate;
            string ReDate;
            string Comments;
            string PP010Doc;
            string ReqCod;
            string ReqNam;
            string YearPdYN;
            string OrderAmt;
            string NegoAmt;
            string TrgtAmt;
            string DrawQty;
            string ReWkRn;
            string Status;

            string BaseEntry;
            string BaseLine;
            string DocType;
            string CurDocDate;

            short MinusNum = 0; //화면모드에 따라 VisualRowCount에서 빼줄 수

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    MinusNum = 2;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    MinusNum = 1;
                }

                oMat01.FlushToDataSource();
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - MinusNum; loopCount++)
                {
                    DocEntry = oDS_PS_PP020H.GetValue("DocEntry", loopCount).ToString().Trim(); //DocEntry는 프로시저에서 처리
                    WorkGbn = oDS_PS_PP020H.GetValue("U_WorkGbn", loopCount).ToString().Trim();
                    JakName = oDS_PS_PP020H.GetValue("U_JakName", loopCount).ToString().Trim();
                    SubNo1 = oDS_PS_PP020H.GetValue("U_SubNo1", loopCount).ToString().Trim();
                    SubNo2 = oDS_PS_PP020H.GetValue("U_SubNo2", loopCount).ToString().Trim();
                    JakMyung = oDS_PS_PP020H.GetValue("U_JakMyung", loopCount).ToString().Trim();
                    JakSize = oDS_PS_PP020H.GetValue("U_JakSize", loopCount).ToString().Trim();
                    JakUnit = oDS_PS_PP020H.GetValue("U_JakUnit", loopCount).ToString().Trim();
                    RegNum = oDS_PS_PP020H.GetValue("U_RegNum", loopCount).ToString().Trim();
                    ItemCode = oDS_PS_PP020H.GetValue("U_ItemCode", loopCount).ToString().Trim();
                    ItemName = dataHelpClass.Make_ItemName(oDS_PS_PP020H.GetValue("U_ItemName", loopCount)).Trim();
                    Material = oDS_PS_PP020H.GetValue("U_Material", loopCount).ToString().Trim();
                    Unit = oDS_PS_PP020H.GetValue("U_Unit", loopCount).ToString().Trim();
                    Size = oDS_PS_PP020H.GetValue("U_Size", loopCount).ToString().Trim();
                    ReqCod = oDS_PS_PP020H.GetValue("U_ReqCod", loopCount).ToString().Trim();
                    ReqNam = oDS_PS_PP020H.GetValue("U_ReqNam", loopCount).ToString().Trim();
                    ItmBsort = oDS_PS_PP020H.GetValue("U_ItmBSort", loopCount).ToString().Trim();
                    SjDocNum = oDS_PS_PP020H.GetValue("U_SjDocNum", loopCount).ToString().Trim();
                    SjLinNum = oDS_PS_PP020H.GetValue("U_SjLinNum", loopCount).ToString().Trim();
                    SjWeight = oDS_PS_PP020H.GetValue("U_SjWeight", loopCount).ToString().Trim();
                    WrWeight = oDS_PS_PP020H.GetValue("U_WrWeight", loopCount).ToString().Trim();
                    SjDcDate = oDS_PS_PP020H.GetValue("U_SjDcDate", loopCount).ToString().Trim();
                    SjDuDate = oDS_PS_PP020H.GetValue("U_SjDuDate", loopCount).ToString().Trim();
                    SlePrice = oDS_PS_PP020H.GetValue("U_SlePrice", loopCount).ToString().Trim();
                    CardCode = oDS_PS_PP020H.GetValue("U_CardCode", loopCount).ToString().Trim();
                    CardName = oDS_PS_PP020H.GetValue("U_CardName", loopCount).ToString().Trim();
                    ShipCode = oDS_PS_PP020H.GetValue("U_ShipCode", loopCount).ToString().Trim();
                    ShipName = oDS_PS_PP020H.GetValue("U_ShipName", loopCount).ToString().Trim();
                    Comments = oDS_PS_PP020H.GetValue("U_Comments", loopCount).ToString().Trim();
                    InOutGbn = oDS_PS_PP020H.GetValue("U_InOutGbn", loopCount).ToString().Trim();
                    JakDate = oDS_PS_PP020H.GetValue("U_JakDate", loopCount).ToString().Trim();
                    ProDate = oDS_PS_PP020H.GetValue("U_ProDate", loopCount).ToString().Trim();
                    ReDate = oDS_PS_PP020H.GetValue("U_ReDate", loopCount).ToString().Trim();
                    PP010Doc = oDS_PS_PP020H.GetValue("U_PP010Doc", loopCount).ToString().Trim();
                    DrawQty = oDS_PS_PP020H.GetValue("U_DrawQty", loopCount).ToString().Trim();
                    YearPdYN = oDS_PS_PP020H.GetValue("U_YearPdYN", loopCount).ToString().Trim();
                    OrderAmt = oDS_PS_PP020H.GetValue("U_OrderAmt", loopCount).ToString().Trim();
                    NegoAmt = oDS_PS_PP020H.GetValue("U_NegoAmt", loopCount).ToString().Trim();
                    TrgtAmt = oDS_PS_PP020H.GetValue("U_TrgtAmt", loopCount).ToString().Trim();
                    ReWkRn = oDS_PS_PP020H.GetValue("U_ReWkRn", loopCount).ToString().Trim();
                    Status = "O";

                    Query01 = "EXEC PS_PP020_03 '";
                    Query01 += DocEntry + "','";
                    Query01 += BPLId + "','";
                    Query01 += JakName + "','";
                    Query01 += SubNo1 + "','";
                    Query01 += SubNo2 + "','";
                    Query01 += JakMyung + "','";
                    Query01 += JakSize + "','";
                    Query01 += JakUnit + "','";
                    Query01 += RegNum + "','";
                    Query01 += ItemCode + "','";
                    Query01 += ItemName + "','";
                    Query01 += Material + "','";
                    Query01 += Unit + "','";
                    Query01 += Size + "','";
                    Query01 += ItmBsort + "','";
                    Query01 += SjDocNum + "','";
                    Query01 += SjLinNum + "','";
                    Query01 += SjWeight + "','";
                    Query01 += WrWeight + "','";
                    Query01 += SjDcDate + "','";
                    Query01 += SjDuDate + "','";
                    Query01 += SlePrice + "','";
                    Query01 += WorkGbn + "','";
                    Query01 += CardCode + "','";
                    Query01 += CardName + "','";
                    Query01 += ShipCode + "','";
                    Query01 += ShipName + "','";
                    Query01 += InOutGbn + "','";
                    Query01 += JakDate + "','";
                    Query01 += ProDate + "','";
                    Query01 += ReDate + "','";
                    Query01 += Comments + "','";
                    Query01 += PP010Doc + "','";
                    Query01 += ReqCod + "','";
                    Query01 += ReqNam + "','";
                    Query01 += YearPdYN + "','";
                    Query01 += OrderAmt + "','";
                    Query01 += NegoAmt + "','";
                    Query01 += TrgtAmt + "','";
                    Query01 += DrawQty + "','";
                    Query01 += ReWkRn + "','";
                    Query01 += Status + "'";

                    //선행프로세스 대비 일자체크_S
                    BaseEntry = PP010Doc;
                    BaseLine = "0";
                    DocType = "PS_PP020";
                    CurDocDate = JakDate;

                    Query02 = "EXEC PS_Z_CHECK_DATE '";
                    Query02 += BaseEntry + "','";
                    Query02 += BaseLine + "','";
                    Query02 += DocType + "','";
                    Query02 += CurDocDate + "'";

                    RecordSet02.DoQuery(Query02);
                    //선행프로세스 대비 일자체크_E

                    if (RecordSet02.Fields.Item("ReturnValue").Value == "True")
                    {
                        RecordSet01.DoQuery(Query01); //등록
                    }
                    else
                    {
                        ErrOrdNum = ErrOrdNum + " [" + ItemCode + "]";
                    }
                }

                //하나라도 선행프로세스 일자가 빠른 작번이 있으면
                if (!string.IsNullOrEmpty(ErrOrdNum))
                {
                    errMessage = "작번등록일은 생산의뢰접수일과 같거나 늦어야합니다. 확인하십시오." + (char)13 + ErrOrdNum;

                    //등록되지 않은 작번이 있어도 화면 Clear_S
                    oMat01.Clear();
                    oMat01.FlushToDataSource();
                    oMat01.LoadFromDataSource();
                    PS_PP020_AddMatrixRow(0, true, "");
                    //등록되지 않은 작번이 있어도 화면 Clear_E

                    throw new Exception();
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                if (returnValue == true)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("처리 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Sub 작번 등록
        /// </summary>
        /// <returns></returns>
        private bool PS_PP020_AddSubJakNum()
        {
            bool returnValue = false;

            short i;
            string sQry;
            string errMessage = string.Empty;
            string SjDocNum;
            string SjLinNum;
            string ItmBsort;
            string Unit;
            string ItemName;
            string RegNum;
            string DocEntry = string.Empty;
            string BPLId;
            string ItemCode;
            string Material;
            string Size;
            double SjWeight;
            double SlePrice;
            string ShipCode;
            string CardCode;
            string SjDuDate;
            string SjDcDate;
            string WorkGbn;
            string CardName;
            string ShipName;
            string JakSize;
            string ReDate;
            string JakDate;
            string SubNo1;
            string Comments;
            string JakName;
            string SubNo2;
            string InOutGbn;
            string ProDate;
            string JakMyung;
            string JakUnit;
            double WrWeight;
            string Status;
            string ReWkRn; //재작업 사유
            string ReWork; //재작업 구분
            double OrderAmt; //수주금액
            double NegoAmt; //Nego금액
            double TrgtAmt; //목표금액(생산)
            string reqCode; //등록자 사번
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                oMat01.FlushToDataSource();

                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

                for (i = 0; i <= oMat02.RowCount - 1; i++)
                {
                    JakName = oMat02.Columns.Item("JakName").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    SubNo1 = oMat02.Columns.Item("SubNo1").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    SubNo2 = oMat02.Columns.Item("SubNo2").Cells.Item(i + 1).Specific.Value.ToString().Trim();

                    ReWkRn = oMat02.Columns.Item("ReWkRn").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    ReWork = oForm.Items.Item("ReWork").Specific.Value;

                    sQry = "Select COUNT(*) From [@PS_PP020H] Where U_JakName = '" + JakName + "' And U_SubNo1 = '" + SubNo1 + "' And U_SubNo2 = '" + SubNo2 + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (Convert.ToInt16(oRecordSet01.Fields.Item(0).Value) > 0)
                    {
                        errMessage = (i + 1) + "번째 라인의 작번이 이미 등록되어 있습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else
                    {
                        if (ReWork == "20" && string.IsNullOrEmpty(ReWkRn))
                        {
                            errMessage = "재작업인 경우는 재작업 사유를 필수로 입력해야합니다.  [" + (i + 1) + "행]";
                            throw new Exception();
                        }
                    }
                }

                reqCode = dataHelpClass.User_MSTCOD();

                for (i = 0; i <= oMat02.RowCount - 1; i++)
                {
                    WorkGbn = oMat02.Columns.Item("WorkGbn").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    JakName = oMat02.Columns.Item("JakName").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    SubNo1 = oMat02.Columns.Item("SubNo1").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    SubNo2 = oMat02.Columns.Item("SubNo2").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    JakMyung = oMat02.Columns.Item("JakMyung").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    JakSize = oMat02.Columns.Item("JakSize").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    JakUnit = oMat02.Columns.Item("JakUnit").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    RegNum = oMat02.Columns.Item("RegNum").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    ItemCode = oMat02.Columns.Item("ItemCode").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    ItemName = dataHelpClass.Make_ItemName(oMat02.Columns.Item("ItemName").Cells.Item(i + 1).Specific.Value.ToString().Trim()).Trim();
                    Material = oMat02.Columns.Item("Material").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    Unit = oMat02.Columns.Item("Unit").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    Size = oMat02.Columns.Item("Size").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    ItmBsort = oMat02.Columns.Item("ItmBSort").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    SjDocNum = oMat02.Columns.Item("SjDocNum").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    SjLinNum = oMat02.Columns.Item("SjLinNum").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    SjWeight = Convert.ToDouble(oMat02.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                    WrWeight = Convert.ToDouble(oMat02.Columns.Item("WrWeight").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                    SjDcDate = oMat02.Columns.Item("SjDcDate").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    SjDuDate = oMat02.Columns.Item("SjDuDate").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    SlePrice = Convert.ToDouble(oMat02.Columns.Item("SlePrice").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                    CardCode = oMat02.Columns.Item("CardCode").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    CardName = oMat02.Columns.Item("CardName").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    ShipCode = oMat02.Columns.Item("ShipCode").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    ShipName = oMat02.Columns.Item("ShipName").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    Comments = oMat02.Columns.Item("Comments").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    InOutGbn = oMat02.Columns.Item("InOutGbn").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    JakDate = oMat02.Columns.Item("JakDate").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    ProDate = oMat02.Columns.Item("ProDate").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    ReDate = oMat02.Columns.Item("ReDate").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    ReWkRn = oMat02.Columns.Item("ReWkRn").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                    OrderAmt = Convert.ToDouble(oMat02.Columns.Item("OrderAmt").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                    NegoAmt = Convert.ToDouble(oMat02.Columns.Item("NegoAmt").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                    TrgtAmt = Convert.ToDouble(oMat02.Columns.Item("TrgtAmt").Cells.Item(i + 1).Specific.Value.ToString().Trim());
                    Status = "O";

                    sQry = "EXEC PS_PP020_04 '";
                    sQry += DocEntry + "','";
                    sQry += BPLId + "','";
                    sQry += JakName + "','";
                    sQry += SubNo1 + "','";
                    sQry += SubNo2 + "','";
                    sQry += JakMyung + "','";
                    sQry += JakSize + "','";
                    sQry += JakUnit + "','";
                    sQry += RegNum + "','";
                    sQry += ItemCode + "','";
                    sQry += ItemName + "','";
                    sQry += Material + "','";
                    sQry += Unit + "','";
                    sQry += Size + "','";
                    sQry += ItmBsort + "','";
                    sQry += SjDocNum + "','";
                    sQry += SjLinNum + "','";
                    sQry += SjWeight + "','";
                    sQry += WrWeight + "','";
                    sQry += SjDcDate + "','";
                    sQry += SjDuDate + "','";
                    sQry += SlePrice + "','";
                    sQry += WorkGbn + "','";
                    sQry += CardCode + "','";
                    sQry += CardName + "','";
                    sQry += ShipCode + "','";
                    sQry += ShipName + "','";
                    sQry += InOutGbn + "','";
                    sQry += JakDate + "','";
                    sQry += ProDate + "','";
                    sQry += ReDate + "','";
                    sQry += Comments + "','";
                    sQry += OrderAmt + "','";
                    sQry += NegoAmt + "','";
                    sQry += TrgtAmt + "','";
                    sQry += Status + "','";
                    sQry += reqCode + "','";
                    sQry += ReWkRn + "'";

                    oRecordSet01.DoQuery(sQry); //등록
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                if (returnValue == true)
                {
                    oDS_PS_TEMPTABLE.Clear();
                    oMat02.Clear();
                    PSH_Globals.SBO_Application.StatusBar.SetText("Sub작번등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Header 필수 입력 사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP020_CheckHeaderDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
                {
                    errMessage = "사업장은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Line 필수 입력 사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP020_CheckLineDataValid()
        {
            bool returnValue = false;
            int i;
            string errMessage = string.Empty;
            int minusLineCount;

            try
            {
                minusLineCount = oForm.Mode == BoFormMode.fm_ADD_MODE ? 1 : 0; //추가모드일 때는 마지막행이 빈행으로 추가되므로 1행을 뺀 만큼 검사하기 위한 변수

                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - minusLineCount; i++)
                {
                    if (oDS_PS_PP020H.GetValue("U_InOutGbn", i - 1).ToString().Trim() == "%")
                    {
                        errMessage = i + "번 라인의 외주구분이 선택되지 않았습니다. 확인하세요.";
                        throw new Exception();
                    }
                }
                oMat01.LoadFromDataSource();

                oMat02.FlushToDataSource();
                for (i = 1; i <= oMat02.VisualRowCount; i++)
                {
                    if (oDS_PS_TEMPTABLE.GetValue("U_sField17", i - 1).ToString().Trim() == "%")
                    {
                        errMessage = "서브작번 " + i + "번 라인의 외주구분이 선택되지 않았습니다. 확인하세요.";
                        throw new Exception();
                    }
                }
                oMat02.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
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
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP020_CheckHeaderDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_PP020_CheckLineDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_PP020_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_PP020_AddMatrixRow(0, true, "");
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP020_CheckLineDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_PP020_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PS_PP020_LoadCaption();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        if (PS_PP020_CheckHeaderDataValid() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PS_PP020_LoadCaption();
                        PS_PP020_LoadData();
                    }
                    else if (pVal.ItemUID == "Btn03") //서브작번등록
                    {
                        if (PS_PP020_CheckLineDataValid() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }

                        if (PS_PP020_AddSubJakNum() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
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
                oForm.Freeze(false);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                            {
                                PS_SM010 tempForm = new PS_SM010();
                                tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "JakName")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("JakName").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "ReqCod")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ReqCod").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "CardCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "ShipCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ShipCode").Cells.Item(pVal.Row).Specific.Value))
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
                if (pVal.Before_Action == true)
                {
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
                    if (pVal.ItemUID == "ReWork")
                    {
                        oDS_PS_TEMPTABLE.Clear();
                        oMat02.Clear();
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PS_PP020_LoadCaption();
                        }
                    }

                    if (pVal.ItemUID == "WorkGbn" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oForm.Freeze(true);
                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();

                        PS_PP020_AddMatrixRow(0, true, "");
                        if (oForm.Items.Item("WorkGbn").Specific.Selected.Value == "10")
                        {
                            oMat01.Columns.Item("RegNum").Editable = true;
                            oMat01.Columns.Item("ItemCode").Editable = false;
                        }
                        else
                        {
                            oMat01.Columns.Item("RegNum").Editable = false;
                            oMat01.Columns.Item("ItemCode").Editable = true;
                        }
                        oForm.Freeze(false);
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
                    if (pVal.ItemUID == "RadioMat01")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat01";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oForm.Freeze(false);
                    }
                    else if (pVal.ItemUID == "RadioMat02")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat02";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oForm.Freeze(false);
                    }

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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            int j;
            string JakName;
            string sQry;
            string SubNo1 = string.Empty;
            string SubNo2 = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.Row != 0)
                    {
                        j = 0;

                        if (oMat02.VisualRowCount == 0)
                        {
                            oDS_PS_TEMPTABLE.Clear();
                        }

                        JakName = oDS_PS_PP020H.GetValue("U_JakName", pVal.Row - 1).ToString().Trim();

                        if (oDS_PS_PP020H.GetValue("U_SubNo1", pVal.Row - 1).ToString().Trim() == "00")
                        {
                            if (oForm.Items.Item("ReWork").Specific.Value.ToString().Trim() == "10") //정상
                            {
                                if (oDS_PS_PP020H.GetValue("U_SubNo2", pVal.Row - 1).ToString().Trim() == "000")
                                {
                                    sQry = "  SELECT  MAX(ISNULL(U_SubNo1, '00')) ";
                                    sQry += " FROM    [@PS_PP020H] ";
                                    sQry += " WHERE   U_JakName = '" + JakName + "'";
                                    sQry += "         AND U_SubNo2 = '" + oDS_PS_PP020H.GetValue("U_SubNo2", pVal.Row - 1).ToString().Trim() + "'";
                                    oRecordSet01.DoQuery(sQry);

                                    SubNo1 = Convert.ToString(Convert.ToInt16(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1).PadLeft(2, '0');
                                    SubNo2 = "000";
                                }
                                else
                                {
                                    PSH_Globals.SBO_Application.MessageBox("해당 작번은 서브작번을 만들 수 없습니다. 확인하세요.");
                                    j = 1;
                                }
                            }
                            else
                            {
                                if (oDS_PS_PP020H.GetValue("U_SubNo2", pVal.Row - 1).ToString().Trim() == "000")
                                {
                                    sQry = "  SELECT  MAX(ISNULL(U_SubNo1, '00')) ";
                                    sQry += " FROM    [@PS_PP020H] ";
                                    sQry += " WHERE   U_JakName = '" + JakName + "' ";
                                    sQry += "         AND U_SubNo1 >= '80' ";
                                    sQry += "         AND U_SubNo2 = '" + oDS_PS_PP020H.GetValue("U_SubNo2", pVal.Row - 1).ToString().Trim() + "'";
                                    oRecordSet01.DoQuery(sQry);

                                    if (string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value.ToString().Trim()))
                                    {
                                        SubNo1 = "80";
                                    }
                                    else
                                    {
                                        SubNo1 = Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1).PadLeft(2, '0');
                                    }
                                    SubNo2 = "000";
                                }
                                else
                                {
                                    PSH_Globals.SBO_Application.MessageBox("해당 작번은 서브작번을 만들 수 없습니다. 확인하세요.");
                                    j = 1;
                                }
                            }
                        }
                        else if (oDS_PS_PP020H.GetValue("U_SubNo1", pVal.Row - 1).ToString().Trim() != "00")
                        {
                            if (oDS_PS_PP020H.GetValue("U_SubNo2", pVal.Row - 1).ToString().Trim() == "000")
                            {
                                sQry = "  SELECT  MAX(ISNULL(U_SubNo2, '000')) ";
                                sQry += " FROM    [@PS_PP020H] Where U_JakName = '" + JakName + "' ";
                                sQry += "         AND U_SubNo1 = '" + oDS_PS_PP020H.GetValue("U_SubNo1", pVal.Row - 1).ToString().Trim() + "'";

                                oRecordSet01.DoQuery(sQry);

                                SubNo1 = oDS_PS_PP020H.GetValue("U_SubNo1", pVal.Row - 1).ToString().Trim();
                                SubNo2 = Convert.ToString(Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1).PadLeft(3, '0');
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.MessageBox("해당 작번은 서브작번을 만들 수 없습니다. 확인하세요.");
                                j = 1;
                            }
                        }

                        for (i = 0; i <= oMat02.VisualRowCount - 1; i++)
                        {
                            oMat02.LoadFromDataSource();
                            if (oDS_PS_TEMPTABLE.GetValue("U_sField01", i).ToString().Trim() == JakName && oDS_PS_TEMPTABLE.GetValue("U_sField02", i).ToString().Trim() == SubNo1 && oDS_PS_TEMPTABLE.GetValue("U_sField03", i).ToString().Trim() == SubNo2)
                            {
                                PSH_Globals.SBO_Application.MessageBox("같은 행을 두번 선택할 수 없습니다. 확인하세요.");
                                j = 1;
                            }
                        }

                        if (j == 0)
                        {
                            oForm.Freeze(true);
                            PS_PP020_AddMatrixRow(oMat02.VisualRowCount, false, "Mat02");
                            oMat02.Columns.Item("JakName").Cells.Item(oMat02.VisualRowCount).Specific.Value = JakName;
                            oMat02.Columns.Item("SubNo1").Cells.Item(oMat02.VisualRowCount).Specific.Value = SubNo1.PadLeft(2, '0');
                            oMat02.Columns.Item("SubNo2").Cells.Item(oMat02.VisualRowCount).Specific.Value = SubNo2.PadLeft(3, '0');
                            oMat02.Columns.Item("ItemCode").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("ItemName").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_ItemName", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("Material").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_Material", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("Unit").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_Unit", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("Size").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_Size", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("ItmBSort").Cells.Item(oMat02.VisualRowCount).Specific.Select(oDS_PS_PP020H.GetValue("U_ItmBSort", pVal.Row - 1).ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat02.Columns.Item("SjDocNum").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_SjDocNum", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("SjLinNum").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_SjLinNum", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("CardCode").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_CardCode", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("CardName").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_CardName", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("ShipCode").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_ShipCode", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("ShipName").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_ShipName", pVal.Row - 1).ToString().Trim();

                            if (!string.IsNullOrEmpty(oDS_PS_PP020H.GetValue("U_InOutGbn", pVal.Row - 1).ToString().Trim()))
                            {
                                oMat02.Columns.Item("InOutGbn").Cells.Item(oMat02.VisualRowCount).Specific.Select(oDS_PS_PP020H.GetValue("U_InOutGbn", pVal.Row - 1).ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }
                            oMat02.Columns.Item("JakDate").Cells.Item(oMat02.VisualRowCount).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat02.Columns.Item("ProDate").Cells.Item(oMat02.VisualRowCount).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat02.Columns.Item("ReDate").Cells.Item(oMat02.VisualRowCount).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat02.Columns.Item("SjDcDate").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_SjDcDate", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("SjDuDate").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_SjDuDate", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("SlePrice").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_SlePrice", pVal.Row - 1).ToString().Trim();
                            oMat02.Columns.Item("WorkGbn").Cells.Item(oMat02.VisualRowCount).Specific.Select(oDS_PS_PP020H.GetValue("U_WorkGbn", pVal.Row - 1).ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat02.Columns.Item("OrderAmt").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_OrderAmt", pVal.Row - 1).ToString().Trim(); //수주금액
                            oMat02.Columns.Item("NegoAmt").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_NegoAmt", pVal.Row - 1).ToString().Trim(); //Nego금액
                            oMat02.Columns.Item("TrgtAmt").Cells.Item(oMat02.VisualRowCount).Specific.Value = oDS_PS_PP020H.GetValue("U_TrgtAmt", pVal.Row - 1).ToString().Trim(); //목표금액

                            oMat02.FlushToDataSource();
                            oMat02.LoadFromDataSource();
                            oMat02.AutoResizeColumns();

                            oForm.Freeze(false);
                            j = 0;
                        }

                        BubbleEvent = false;
                    }
                    else if (pVal.ItemUID == "Mat01" && pVal.Row == 0 && pVal.ColUID == "InOutGbn") //외주구분의 첫행을 선택한 후 컬럼 타이틀을 더블클릭하면 첫행의 값으로 첫행 이외의 값을 자동으로 선택하는 기능
                    {
                        oForm.Freeze(true);

                        oMat01.FlushToDataSource();
                        string FirstInOutGbn = oDS_PS_PP020H.GetValue("U_InOutGbn", 0).ToString().Trim();
                        for (i = 1; i <= oMat01.VisualRowCount - 2; i++)
                        {
                            oDS_PS_PP020H.SetValue("U_InOutgbn", i, FirstInOutGbn);
                        }
                        oMat01.LoadFromDataSource();
                        oForm.Freeze(false);
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
                        if (pVal.ItemUID == "CntcCode")
                        {
                            PS_PP020_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "CardCode")
                        {
                            PS_PP020_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            PS_PP020_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "JakName")
                            {
                                PS_PP020_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "ItemCode")
                            {
                                PS_PP020_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "ShipCode")
                            {
                                PS_PP020_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "ReqCod")
                            {
                                PS_PP020_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                            }
                            else
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                PS_PP020_LoadCaption();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP020H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_TEMPTABLE);
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
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oForm.Items.Item("Mat01").Top = 82;
                    oForm.Items.Item("Mat01").Left = 6;
                    oForm.Items.Item("Mat01").Width = oForm.Width - 18;
                    oForm.Items.Item("Mat01").Height = (oForm.Height - oForm.Items.Item("Mat01").Top - (oForm.Height - oForm.Items.Item("Btn01").Top)) / 3 * 2 - 25;

                    oForm.Items.Item("RadioMat02").Top = oForm.Items.Item("Mat01").Height + oForm.Items.Item("Mat01").Top - 4;
                    oForm.Items.Item("RadioMat02").Left = 6;
                    oForm.Items.Item("RadioMat02").Height = 20;

                    oForm.Items.Item("33").Top = oForm.Items.Item("RadioMat02").Top;
                    oForm.Items.Item("ReWork").Top = oForm.Items.Item("RadioMat02").Top;

                    oForm.Items.Item("Mat02").Top = oForm.Items.Item("Mat01").Height + oForm.Items.Item("Mat01").Top + 15;
                    oForm.Items.Item("Mat02").Left = oForm.Items.Item("Mat01").Left;
                    oForm.Items.Item("Mat02").Width = oForm.Items.Item("Mat01").Width;
                    oForm.Items.Item("Mat02").Height = oForm.Items.Item("Mat01").Height / 2 + 5;

                    oMat01.AutoResizeColumns();
                    oMat02.AutoResizeColumns();
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            int i;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                            if (oLastItemUID01 == "Mat01")
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("해당 라인의 작번을 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "Yes", "No") == 1)
                                {
                                    //작업지시등록여부 체크
                                    sQry = "  SELECT  COUNT(*) AS [Cnt]";
                                    sQry += " FROM    [@PS_PP030H]";
                                    sQry += " WHERE   U_BaseNum = '" + oLastRightClickDocEntry + "'";
                                    sQry += "         AND [Canceled] = 'N'";
                                    sQry += "         AND U_OrdGbn IN ('105','106')";

                                    oRecordSet.DoQuery(sQry);

                                    //작업지시등록이 존재하면
                                    if (Convert.ToInt16(oRecordSet.Fields.Item("Cnt").Value) > 0)
                                    {
                                        PSH_Globals.SBO_Application.MessageBox("작업지시등록이 존재하는 작번입니다. 삭제할 수 없습니다.");
                                        BubbleEvent = false;
                                        return;
                                    }

                                    sQry = "  DELETE ";
                                    sQry += " FROM    [@PS_PP020H] ";
                                    sQry += " WHERE   DocEntry = '" + oLastRightClickDocEntry + "'";
                                    oRecordSet.DoQuery(sQry);

                                    oLastRightClickDocEntry = "0";
                                    PSH_Globals.SBO_Application.StatusBar.SetText("삭제되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1299": //행닫기
                            if (oLastItemUID01 == "Mat01")
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("해당 라인의 작번등록을 [닫기]처리합니다. 복원할 수 없습니다. 진행하시겠습니까?", 1, "Yes", "No") == 1)
                                {
                                    sQry = "  UPDATE  [@PS_PP020H] ";
                                    sQry += " SET     Status = 'C',";
                                    sQry += "         UpdateDate = GETDATE(),";
                                    sQry += "         UserSign = '" + PSH_Globals.oCompany.UserSignature + "'";
                                    sQry += " Where   DocEntry = '" + oLastRightClickDocEntry + "'";
                                    oRecordSet.DoQuery(sQry);

                                    oLastRightClickDocEntry = "0";
                                    PSH_Globals.SBO_Application.StatusBar.SetText("처리되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
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
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("DocNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PS_PP020H.RemoveRecord(oDS_PS_PP020H.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();
                                oForm.Freeze(false);
                            }

                            if (oMat02.RowCount != oMat02.VisualRowCount - 1)
                            {
                                oForm.Freeze(true);
                                for (i = 0; i <= oMat02.VisualRowCount - 1; i++)
                                {
                                    oMat02.Columns.Item("DocNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat02.FlushToDataSource();
                                oDS_PS_TEMPTABLE.RemoveRecord(oDS_PS_TEMPTABLE.Size - 1);
                                oMat02.Clear();
                                oMat02.LoadFromDataSource();
                                oForm.Freeze(false);
                            }
                            break;
                        case "1281": //찾기
                            PS_PP020_EnableFormItem();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PS_PP020_LoadCaption();
                            break;
                        case "1282": //추가
                            oForm.Freeze(true);
                            PS_PP020_EnableFormItem();

                            if (oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim() == "10")
                            {
                                oMat01.Columns.Item("RegNum").Editable = true;
                                oMat01.Columns.Item("ItemCode").Editable = false;
                            }
                            else
                            {
                                oMat01.Columns.Item("RegNum").Editable = false;
                                oMat01.Columns.Item("ItemCode").Editable = true;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_PP020_AddMatrixRow(0, true, "");
                            PS_PP020_LoadCaption();
                            oForm.Freeze(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_PP020_EnableFormItem();
                            if (oMat01.VisualRowCount > 0)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("CGNo").Cells.Item(oMat01.VisualRowCount).Specific.Value))
                                {
                                    if (oDS_PS_PP020H.GetValue("Status", 0) == "O")
                                    {
                                        PS_PP020_AddMatrixRow(oMat01.RowCount, false, "");
                                    }
                                }
                            }
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
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
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        if (oLastItemUID01 == "Mat01")
                        {
                            oLastRightClickDocEntry = oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
                        }
                    }
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
