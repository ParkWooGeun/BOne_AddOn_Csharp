using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 생산요청
	/// </summary>
	internal class PS_SD010 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SD010H; //등록헤더
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_RightClick_RegNum;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>		
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD010.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD010_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD010");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				PS_SD010_CreateItems();
				PS_SD010_SetComboBox();
				PS_SD010_Initialize();
				PS_SD010_SetRadioButton();
				PS_SD010_AddMatrixRow(0, true);
				PS_SD010_LoadCaption();
				PS_SD010_EnableFormItem();

				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1285", false); //복원
				oForm.EnableMenu("1284", true); //취소
				oForm.EnableMenu("1293", true); //행삭제
				oForm.EnableMenu("1299", true); //행닫기
			}
			catch(Exception ex)
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
        private void PS_SD010_CreateItems()
        {
            try
            {
                oDS_PS_SD010H = oForm.DataSources.DBDataSources.Item("@PS_SD010H");
                oMat01 = oForm.Items.Item("Mat01").Specific; //메트릭스 개체 할당
                oMat01.AutoResizeColumns();

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

                //품목대분류
                oForm.DataSources.UserDataSources.Add("ItmBSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
                oForm.Items.Item("ItmBSort").Specific.DataBind.SetBound(true, "", "ItmBSort");

                //고객
                oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

                //고객명
                oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

                //담당자
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

                //담당자명
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

                //요청일자(시작)
                oForm.DataSources.UserDataSources.Add("RegDateFr", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("RegDateFr").Specific.DataBind.SetBound(true, "", "RegDateFr");

                //요청일자(종료)
                oForm.DataSources.UserDataSources.Add("RegDateTo", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("RegDateTo").Specific.DataBind.SetBound(true, "", "RegDateTo");

                //요청번호(시작)
                oForm.DataSources.UserDataSources.Add("RegNumFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RegNumFr").Specific.DataBind.SetBound(true, "", "RegNumFr");

                //요청번호(종료)
                oForm.DataSources.UserDataSources.Add("RegNumTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("RegNumTo").Specific.DataBind.SetBound(true, "", "RegNumTo");

                //품목코드
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                //품목명
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                //상태
                oForm.DataSources.UserDataSources.Add("Status", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Status").Specific.DataBind.SetBound(true, "", "Status");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD010_SetComboBox()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("BPLId").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //품목대분류
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where U_PudYN = 'Y' Order by Code";
                oRecordSet01.DoQuery(sQry);

                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("ItmBSort").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("ItmBSort").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //상태
                oForm.Items.Item("Status").Specific.ValidValues.Add("O", "열기");
                oForm.Items.Item("Status").Specific.ValidValues.Add("C", "닫기");

                //요청구분
                oMat01.Columns.Item("ReGbn").ValidValues.Add("10", "계획생산요청");
                oMat01.Columns.Item("ReGbn").ValidValues.Add("20", "수주생산요청");
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
        /// 화면 초기화
        /// </summary>
        private void PS_SD010_Initialize()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Radio Button 초기화
        /// </summary>
        private void PS_SD010_SetRadioButton()
        {
            try
            {
                oForm.DataSources.UserDataSources.Add("RadioBtn01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
                oForm.Items.Item("PRType01").Specific.ValOn = "10";
                oForm.Items.Item("PRType01").Specific.ValOff = "0";
                oForm.Items.Item("PRType01").Specific.DataBind.SetBound(true, "", "RadioBtn01");
                oForm.Items.Item("PRType01").Specific.Selected = true;

                oForm.Items.Item("PRType02").Specific.ValOn = "20";
                oForm.Items.Item("PRType02").Specific.ValOff = "0";
                oForm.Items.Item("PRType02").Specific.DataBind.SetBound(true, "", "RadioBtn01");
                oForm.Items.Item("PRType02").Specific.GroupWith("PRType01");

                oMat01.Columns.Item("SjDocLin").Editable = false;
                oMat01.Columns.Item("ItemCode").Editable = true;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD010_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_SD010H.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD010H.Offset = oRow;
                oDS_PS_SD010H.SetValue("DocNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메인버튼 캡션 설정
        /// </summary>
        private void PS_SD010_LoadCaption()
        {
            try
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
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각 모드에 따른 아이템설정
        /// </summary>
        private void PS_SD010_EnableFormItem()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("PRType01").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("ItmBSort").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;

                    oForm.Items.Item("RegDateFr").Enabled = false;
                    oForm.Items.Item("RegDateTo").Enabled = false;
                    oForm.Items.Item("RegNumFr").Enabled = false;
                    oForm.Items.Item("RegNumTo").Enabled = false;
                    oForm.Items.Item("ItemCode").Enabled = false;
                    oForm.Items.Item("Status").Enabled = false;
                    oForm.Items.Item("Btn02").Enabled = false;

                    oMat01.Columns.Item("ItemName").Editable = false;
                    oMat01.Columns.Item("RegNum").Editable = false;
                    oMat01.Columns.Item("SjWeight").Editable = false;
                    oMat01.Columns.Item("UnWeight").Editable = false;
                    oMat01.Columns.Item("Price").Editable = false;
                    oMat01.Columns.Item("LinTotal").Editable = false;
                    oMat01.Columns.Item("BPLId").Editable = false;
                    oMat01.Columns.Item("CntcCode").Editable = false;
                    oMat01.Columns.Item("CntcName").Editable = false;

                    oMat01.Columns.Item("ReGbn").Editable = false;
                    oMat01.Columns.Item("CardCode").Editable = false;
                    oMat01.Columns.Item("CardName").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("PRType01").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("ItmBSort").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;

                    oForm.Items.Item("RegDateFr").Enabled = true;
                    oForm.Items.Item("RegDateTo").Enabled = true;
                    oForm.Items.Item("RegNumFr").Enabled = true;
                    oForm.Items.Item("RegNumTo").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("Status").Enabled = true;
                    oForm.Items.Item("Btn02").Enabled = true;

                    oMat01.Columns.Item("ItemName").Editable = false;
                    oMat01.Columns.Item("RegNum").Editable = false;
                    oMat01.Columns.Item("SjWeight").Editable = false;
                    oMat01.Columns.Item("UnWeight").Editable = false;
                    oMat01.Columns.Item("Price").Editable = false;
                    oMat01.Columns.Item("LinTotal").Editable = false;
                    oMat01.Columns.Item("BPLId").Editable = false;
                    oMat01.Columns.Item("CntcCode").Editable = true;
                    oMat01.Columns.Item("CntcName").Editable = false;

                    oMat01.Columns.Item("ReGbn").Editable = false;
                    oMat01.Columns.Item("CardCode").Editable = false;
                    oMat01.Columns.Item("CardName").Editable = false;
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
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_SD010_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int Qty;
            string sQry;
            string SjDocLin;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                switch (oUID)
                {
                    case "CntcCode":
                        sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("CntcName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "CardCode":
                        sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("CardName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "ItemCode":
                        sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oForm.Items.Item("ItemName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        break;
                    case "Mat01":
                        if (oCol == "SjDocLin")
                        {
                            oForm.Freeze(true);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("SjDocLin").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_SD010_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }

                            SjDocLin = oMat01.Columns.Item("SjDocLin").Cells.Item(oRow).Specific.Value.ToString().Trim();

                            sQry = "  Select  b.ItemCode,";
                            sQry += "         b.Dscription,";
                            sQry += "         c.U_ItmBSort,";
                            sQry += "         b.U_Qty,";
                            sQry += "         b.Quantity,";
                            sQry += "         b.U_UnWeight,";
                            sQry += "         b.Price,";
                            sQry += "         b.LineTotal,";
                            sQry += "         c.U_Size,";
                            sQry += "         c.U_Spec3,";
                            sQry += "         c.U_Quality,";
                            sQry += "         a.BPLId,";
                            sQry += "         b.WhsCode,";
                            sQry += "         d.WhsName,";
                            sQry += "         a.CardCode,";
                            sQry += "         a.CardName,";
                            sQry += "         b.ShipDate,";
                            sQry += "         c.InvntryUom";
                            sQry += " From    [ORDR] a";
                            sQry += "         Inner Join";
                            sQry += "         [RDR1] b";
                            sQry += "             On a.DocEntry = b.DocEntry";
                            sQry += "         Inner Join";
                            sQry += "         [OITM] c";
                            sQry += "             On c.ItemCode = b.ItemCode";
                            sQry += "         Inner Join";
                            sQry += "         [OWHS] d";
                            sQry += "             On d.WhsCode = b.WhsCode";
                            sQry += " Where   a.DocNum = Left('" + SjDocLin + "', CharIndex('-', '" + SjDocLin + "') - 1)";
                            sQry += "         And b.U_LineNum = Right('" + SjDocLin + "', Len('" + SjDocLin + "') - CharIndex('-', '" + SjDocLin + "'))";
                            oRecordSet01.DoQuery(sQry);

                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim();
                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("Dscription").Value.ToString().Trim();
                            oMat01.Columns.Item("RegDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                            oMat01.Columns.Item("DueDate").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("ShipDate").Value.ToString("yyyyMMdd");
                            oMat01.Columns.Item("SjWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim();
                            oMat01.Columns.Item("SjUnit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("InvntryUom").Value.ToString().Trim();
                            oMat01.Columns.Item("UnWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_UnWeight").Value.ToString().Trim();
                            oMat01.Columns.Item("Price").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("Price").Value.ToString().Trim();
                            oMat01.Columns.Item("LinTotal").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("LineTotal").Value.ToString().Trim();
                            oMat01.Columns.Item("BPLId").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("BPLId").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat01.Columns.Item("ReWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim();
                            oMat01.Columns.Item("Unit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("InvntryUom").Value.ToString().Trim();
                            oMat01.Columns.Item("CardCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("CardCode").Value.ToString().Trim();
                            oMat01.Columns.Item("CardName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("CardName").Value.ToString().Trim();

                            oMat01.Columns.Item("DueDate").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                                    PS_SD010_AddMatrixRow(oMat01.RowCount, false);

                                    oMat01.Columns.Item("RegDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                                    oMat01.Columns.Item("DueDate").Cells.Item(oRow).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                                }
                            }

                            sQry = "Select ItemName, U_UnWeight, InvntryUom From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            oMat01.Columns.Item("UnWeight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_UnWeight").Value.ToString().Trim();
                            oMat01.Columns.Item("SjUnit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("InvntryUom").Value.ToString().Trim();
                            oMat01.Columns.Item("Unit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("InvntryUom").Value.ToString().Trim();
                            
                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Freeze(false);
                        }
                        else if (oCol == "ReQty")
                        {
                            oForm.Freeze(true);
                            oMat01.FlushToDataSource();
                            if (!string.IsNullOrEmpty(oDS_PS_SD010H.GetValue("U_ReQty", oRow - 1)))
                            {
                                Qty = Convert.ToInt32(oDS_PS_SD010H.GetValue("U_ReQty", oRow - 1));
                                oDS_PS_SD010H.SetValue("U_ReWeight", oRow - 1, Convert.ToString(Qty));
                            }
                            else
                            {
                                oDS_PS_SD010H.SetValue("U_ReWeight", oRow - 1, "0");
                            }
                            oMat01.LoadFromDataSource();

                            oMat01.Columns.Item("ReQty").Cells.Item(oRow).Click();
                            oForm.Freeze(false);
                        }
                        else if (oCol == "CntcCode")
                        {
                            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oMat01.Columns.Item("CntcCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            oMat01.Columns.Item("CntcName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (oCol == "WhsCode")
                        {
                            sQry = "Select WhsName From [OWHS] Where WhsCode = '" + oMat01.Columns.Item("WhsCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            oMat01.Columns.Item("WhsName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (oCol == "BItemCD")
                        {
                            sQry = "  SELECT  T0.ItemName,";
                            sQry += "         (SELECT TOP 1 U_Minor FROM [@PS_SY001L] WHERE Code = 'M010') AS [WCntcCD],";
                            sQry += "         (SELECT TOP 1 U_CdName FROM [@PS_SY001L] WHERE Code = 'M010') AS [WCntcNM]";
                            sQry += " FROM    [OITM] AS T0";
                            sQry += " WHERE   T0.ItemCode = '" + oMat01.Columns.Item("BItemCD").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            oMat01.Columns.Item("BItemNM").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
                            oMat01.Columns.Item("WCntcCD").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("WCntcCD").Value.ToString().Trim();
                            oMat01.Columns.Item("WCntcNM").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("WCntcNM").Value.ToString().Trim();
                        }
                        else if (oCol == "WCntcCD")
                        {
                            sQry = "  SELECT  T0.U_FullName";
                            sQry += " FROM    [@PH_PY001A] AS T0";
                            sQry += " WHERE   T0.Code = '" + oMat01.Columns.Item("WCntcCD").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            oMat01.Columns.Item("WCntcNM").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_FullName").Value.ToString().Trim();
                        }
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD010_AddData()
        {
            bool returnValue = false;
            short loopCount;
            string ErrNum = null;
            string ErrOrdNum = null;
            string Query01;
            string Query02;
            string SjDocLin; //수주문서-행번호
            string ItemCode; //품목코드
            string ItemName; //품목명
            string RegDate; //요청일자
            string DueDate; //납기
            string RegNum; //요청번호
            string ItmBsort; //품목대분류
            double SjWeight; //수주중량
            string SjUnit; //수주단위
            double UnWeight; //단위중량
            double Price; //단가
            double LinTotal; //금액
            string BPLId; //사업장코드
            string CntcCode; //담당자사번
            string CntcName; //담당자성명
            double ReWeight; //생산요청수량
            string Unit; //단위
            string Comments; //비고
            string ReGbn; //요청구분(계획생산요청, 수주생산요청)
            string CardCode; //고객
            string CardName; //고객명
            string BItemCD; //원소재코드
            string BItemNM; //원소재명
            string WCntcCD; //생산담당사번
            string WCntcNM; //생산담당성명
            string ReqCmt; //구매요청비고
            string Status; //상태
            string BaseEntry;
            string BaseLine;
            string DocType;
            string CurDocDate;
            short MinusNum = 0;

            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                ReGbn = oForm.DataSources.UserDataSources.Item("RadioBtn01").Value.ToString().Trim();
                ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
                CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();

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
                    SjDocLin = oDS_PS_SD010H.GetValue("U_SjDocLin", loopCount).ToString().Trim();
                    ItemCode = oDS_PS_SD010H.GetValue("U_ItemCode", loopCount).ToString().Trim();
                    ItemName = oDS_PS_SD010H.GetValue("U_ItemName", loopCount).ToString().Trim();
                    RegDate = oDS_PS_SD010H.GetValue("U_RegDate", loopCount).ToString().Trim();
                    DueDate = oDS_PS_SD010H.GetValue("U_DueDate", loopCount).ToString().Trim();
                    RegNum = oDS_PS_SD010H.GetValue("U_RegNum", loopCount).ToString().Trim();
                    SjWeight = Convert.ToDouble(oDS_PS_SD010H.GetValue("U_SjWeight", loopCount));
                    SjUnit = oDS_PS_SD010H.GetValue("U_SjUnit", loopCount).ToString().Trim();
                    UnWeight = Convert.ToDouble(oDS_PS_SD010H.GetValue("U_UnWeight", loopCount));
                    Price = Convert.ToDouble(oDS_PS_SD010H.GetValue("U_Price", loopCount));
                    LinTotal = Convert.ToDouble(oDS_PS_SD010H.GetValue("U_LinTotal", loopCount));
                    ReWeight = Convert.ToDouble(oDS_PS_SD010H.GetValue("U_ReWeight", loopCount));
                    Unit = oDS_PS_SD010H.GetValue("U_Unit", loopCount).ToString().Trim();
                    Comments = oDS_PS_SD010H.GetValue("U_Comments", loopCount).ToString().Trim();
                    BItemCD = oDS_PS_SD010H.GetValue("U_BItemCD", loopCount).ToString().Trim();
                    BItemNM = oDS_PS_SD010H.GetValue("U_BItemNM", loopCount).ToString().Trim();
                    WCntcCD = oDS_PS_SD010H.GetValue("U_WCntcCD", loopCount).ToString().Trim();
                    WCntcNM = oDS_PS_SD010H.GetValue("U_WCntcNM", loopCount).ToString().Trim();
                    ReqCmt = oDS_PS_SD010H.GetValue("U_ReqCmt", loopCount).ToString().Trim();
                    Status = "O";

                    Query01 = "EXEC PS_SD010_03 '";
                    Query01 += SjDocLin + "','"; //수주문서-행번호
                    Query01 += ItemCode + "','"; //품목코드
                    Query01 += ItemName + "','"; //품목명
                    Query01 += RegDate + "','"; //요청일자
                    Query01 += DueDate + "','"; //납기
                    Query01 += RegNum + "','"; //요청번호
                    Query01 += ItmBsort + "','"; //품목대분류
                    Query01 += SjWeight + "','"; //수주중량
                    Query01 += SjUnit + "','"; //수주단위
                    Query01 += UnWeight + "','"; //단위중량
                    Query01 += Price + "','"; //단가
                    Query01 += LinTotal + "','"; //금액
                    Query01 += BPLId + "','"; //사업장코드
                    Query01 += CntcCode + "','"; //담당자사번
                    Query01 += CntcName + "','"; //담당자성명
                    Query01 += ReWeight + "','"; //생산요청수량
                    Query01 += Unit + "','"; //단위
                    Query01 += Comments + "','"; //비고
                    Query01 += ReGbn + "','"; //요청구분(계획생산요청, 수주생산요청)
                    Query01 += CardCode + "','"; //고객
                    Query01 += CardName + "','"; //고객명
                    Query01 += BItemCD + "', '"; //원소재코드
                    Query01 += BItemNM + "', '"; //원소재명
                    Query01 += WCntcCD + "', '"; //생산담당사번
                    Query01 += WCntcNM + "', '"; //생산담당성명
                    Query01 += ReqCmt + "', '"; //구매요청비고
                    Query01 += Status + "'"; //상태
                    
                    if (ReGbn == "20") //수주생산요청일 때만 선행프로세스 대비 일자 체크
                    {
                        //선행프로세스 대비 일자체크_S
                        BaseEntry = SjDocLin.Split('-')[0];
                        BaseLine = SjDocLin.Split('-')[1];
                        DocType = "PS_SD010";
                        CurDocDate = RegDate;

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
                    else //계획생산요청일 때는 수주가 없기때문에 선행프로세스 대비 일자 체크 없이 등록
                    {
                        RecordSet01.DoQuery(Query01); //등록
                    }
                }

                //하나라도 선행프로세스 일자가 빠른 작번이 있으면
                if (!string.IsNullOrEmpty(ErrOrdNum))
                {
                    ErrNum = "1";
                    throw new Exception();
                }

                PSH_Globals.SBO_Application.MessageBox("저장 완료!");

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNum == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("생산요청일은 수주일과 같거나 늦어야합니다. 확인하십시오." + (char)13 + ErrOrdNum, 1);

                    //등록되지 않은 작번이 있어도 화면 Clear_S
                    oMat01.Clear();
                    oMat01.FlushToDataSource();
                    oMat01.LoadFromDataSource();
                    PS_SD010_AddMatrixRow(0, true);
                    //등록되지 않은 작번이 있어도 화면 Clear_E
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
            }

            return returnValue;
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PS_SD010_LoadData()
        {
            short i;
            string sQry;
            string ItemCode;
            string RegNumFr;
            string RegDateFr;
            string CardCode;
            string BPLId;
            string ReGbn;
            string ItmBsort;
            string CntcCode;
            string RegDateTo;
            string RegNumTo;
            string Status;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            { 
                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                ReGbn = oForm.DataSources.UserDataSources.Item("RadioBtn01").Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
                RegDateFr = oForm.Items.Item("RegDateFr").Specific.Value.ToString().Trim();
                RegDateTo = oForm.Items.Item("RegDateTo").Specific.Value.ToString().Trim();
                RegNumFr = oForm.Items.Item("RegNumFr").Specific.Value.ToString().Trim();
                RegNumTo = oForm.Items.Item("RegNumTo").Specific.Value.ToString().Trim();
                ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                Status = oForm.Items.Item("Status").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(BPLId))
                {
                    BPLId = "%";
                }
                if (string.IsNullOrEmpty(ItmBsort) || ItmBsort == "ALL")
                {
                    ItmBsort = "%";
                }
                if (string.IsNullOrEmpty(CardCode))
                {
                    CardCode = "%";
                }   
                if (string.IsNullOrEmpty(CntcCode))
                {
                    CntcCode = "%";
                }   
                if (string.IsNullOrEmpty(RegDateFr))
                {
                    RegDateFr = "19000101";
                }
                if (string.IsNullOrEmpty(RegDateTo))
                {
                    RegDateTo = "20991231";
                }
                if (string.IsNullOrEmpty(RegNumFr))
                {
                    RegNumFr = "0000000000";
                }
                if (string.IsNullOrEmpty(RegNumTo))
                {
                    RegNumTo = "9999999999";
                }
                if (string.IsNullOrEmpty(ItemCode))
                {
                    ItemCode = "%";
                }
                if (string.IsNullOrEmpty(Status))
                {
                    Status = "%";
                }

                sQry = "EXEC [PS_SD010_01] '";
                sQry += ReGbn + "','";
                sQry += BPLId + "','";
                sQry += ItmBsort + "','";
                sQry += CardCode + "','";
                sQry += CntcCode + "','";
                sQry += RegDateFr + "','";
                sQry += RegDateTo + "','";
                sQry += RegNumFr + "','";
                sQry += RegNumTo + "','";
                sQry += ItemCode + "','";
                sQry += Status + "'";

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_SD010H.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "조회 결과가 없습니다.확인하세요.";
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_SD010H.Size)
                    {
                        oDS_PS_SD010H.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_SD010H.Offset = i;
                    oDS_PS_SD010H.SetValue("DocNum", i, Convert.ToString(i + 1));
                    oDS_PS_SD010H.SetValue("U_SjDocLin", i, oRecordSet01.Fields.Item("U_SjDocLin").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_RegDate", i, oRecordSet01.Fields.Item("U_RegDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_SD010H.SetValue("U_DueDate", i, oRecordSet01.Fields.Item("U_DueDate").Value.ToString("yyyyMMdd"));
                    oDS_PS_SD010H.SetValue("U_RegNum", i, oRecordSet01.Fields.Item("U_RegNum").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_ItmBSort", i, oRecordSet01.Fields.Item("U_ItmBSort").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_SjWeight", i, oRecordSet01.Fields.Item("U_SjWeight").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_UnWeight", i, oRecordSet01.Fields.Item("U_UnWeight").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_Price", i, oRecordSet01.Fields.Item("U_Price").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_LinTotal", i, oRecordSet01.Fields.Item("U_LinTotal").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_BPLId", i, oRecordSet01.Fields.Item("U_BPLId").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_CntcCode", i, oRecordSet01.Fields.Item("U_CntcCode").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_CntcName", i, oRecordSet01.Fields.Item("U_CntcName").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_ReWeight", i, oRecordSet01.Fields.Item("U_ReWeight").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_Comments", i, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_ReGbn", i, oRecordSet01.Fields.Item("U_ReGbn").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_CardCode", i, oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_CardName", i, oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_BItemCD", i, oRecordSet01.Fields.Item("U_BItemCD").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_BItemNM", i, oRecordSet01.Fields.Item("U_BItemNM").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_WCntcCD", i, oRecordSet01.Fields.Item("U_WCntcCD").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_WCntcNM", i, oRecordSet01.Fields.Item("U_WCntcNM").Value.ToString().Trim());
                    oDS_PS_SD010H.SetValue("U_Status", i, oRecordSet01.Fields.Item("U_Status").Value.ToString().Trim());

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
        /// 선행프로세스와 일자 비교(구현만 하고 사용은 안함)
        /// </summary>
        /// <returns>true-선행프로세스보다 일자가 같거나 느릴 경우, false-선행프로세스보다 일자가 빠를 경우</returns>
        private bool PS_SD010_CheckDate()
        {
            bool returnValue = false;
            string Query01;
            short loopCount;
            string BaseEntry;
            string BaseLine;
            string DocType;
            string CurDocDate;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 2; loopCount++)
                {
                    BaseEntry = oMat01.Columns.Item("SjDocLin").Cells.Item(loopCount).Specific.Value.Split('-')[0];
                    BaseLine = oMat01.Columns.Item("SjDocLin").Cells.Item(loopCount).Specific.Value.Split('-')[1];
                    DocType = "PS_SD010";
                    CurDocDate = oMat01.Columns.Item("RegDate").Cells.Item(loopCount).Specific.Value;

                    Query01 = "EXEC PS_Z_CHECK_DATE '";
                    Query01 += BaseEntry + "','";
                    Query01 += BaseLine + "','";
                    Query01 += DocType + "','";
                    Query01 += CurDocDate + "'";

                    oRecordSet01.DoQuery(Query01);

                    if (oRecordSet01.Fields.Item("ReturnValue").Value == "True")
                    {
                        returnValue = true;
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 필수입력사항 체크(헤더)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD010_DeleteHeaderSpaceLine()
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
                else if (string.IsNullOrEmpty(oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim()))
                {
                    errMessage = "품목대분류는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
                    {
                        errMessage = "담당자는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (oForm.DataSources.UserDataSources.Item("RadioBtn01").Value == "20")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
                        {
                            errMessage = "수주생산요청시 고객은 필수사항입니다. 확인하세요.";
                            throw new Exception();
                        }
                    }
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
        /// 필수입력사항 체크(라인)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD010_DeleteMatrixSpaceLine()
        {
            bool returnValue = false;
            int i;
            int j;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1 && string.IsNullOrEmpty(oDS_PS_SD010H.GetValue("U_ItemCode", 0).ToString().Trim()))
                {
                    j = 0;
                }
                else
                {
                    j = oMat01.VisualRowCount;
                }

                for (i = 0; i <= j - 2; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_SD010H.GetValue("U_ItemCode", i)))
                    {
                        errMessage = (i + 1) + "번 라인의 품목코드가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (Convert.ToDouble(oDS_PS_SD010H.GetValue("U_ReWeight", i)) == 0)
                    {
                        errMessage = (i + 1) + "번 라인의 생산요청수량이 0입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_SD010H.GetValue("U_RegDate", i)))
                    {
                        errMessage = (i + 1) + "번 라인의 생산요청일자가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_SD010H.GetValue("U_DueDate", i)))
                    {
                        errMessage = (i + 1) + "번 라인의 납기일자가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (!string.IsNullOrEmpty(oDS_PS_SD010H.GetValue("U_BItemCD", i)))
                    {
                        if (string.IsNullOrEmpty(oDS_PS_SD010H.GetValue("U_WCntcCD", i)))
                        {
                            errMessage = (i + 1) + "번 라인의 생산담당자 정보가 없습니다." + (char)13 + " 원소재를 등록하는경우 생산담당자가 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_SD010H.GetValue("U_ReqCmt", i)))
                        {
                            errMessage = (i + 1) + "번 라인의 구매요청비고 정보가 없습니다." + (char)13 + "원소재를 등록하는경우 구매요청비고가 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                    }
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

            return returnValue;
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        [STAThread]
        private void PS_SD010_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string sQry01;
            string ItemCode;
            string RegNumFr;
            string RegDateFr;
            string CardCode;
            string BPLId;
            string ReGbn;
            string ItmBsort;
            string CntcCode;
            string RegDateTo;
            string RegNumTo;
            string Status;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ReGbn = oForm.DataSources.UserDataSources.Item("RadioBtn01").Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
                RegDateFr = oForm.Items.Item("RegDateFr").Specific.Value.ToString().Trim();
                RegDateTo = oForm.Items.Item("RegDateTo").Specific.Value.ToString().Trim();
                RegNumFr = oForm.Items.Item("RegNumFr").Specific.Value.ToString().Trim();
                RegNumTo = oForm.Items.Item("RegNumTo").Specific.Value.ToString().Trim();
                ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                Status = oForm.Items.Item("Status").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(BPLId))
                {
                    BPLId = "%";
                }
                if (string.IsNullOrEmpty(ItmBsort) || ItmBsort == "ALL")
                {
                    ItmBsort = "%";
                }
                if (string.IsNullOrEmpty(CardCode))
                {
                    CardCode = "%";
                }
                if (string.IsNullOrEmpty(CntcCode))
                {
                    CntcCode = "%";
                }
                if (string.IsNullOrEmpty(RegDateFr))
                {
                    RegDateFr = "19000101";
                }
                if (string.IsNullOrEmpty(RegDateTo))
                {
                    RegDateTo = "20991231";
                }
                if (string.IsNullOrEmpty(RegNumFr))
                {
                    RegNumFr = "0000000000";
                }
                if (string.IsNullOrEmpty(RegNumTo))
                {
                    RegNumTo = "9999999999";
                }
                if (string.IsNullOrEmpty(ItemCode))
                {
                    ItemCode = "%";
                }
                if (string.IsNullOrEmpty(Status))
                {
                    Status = "%";
                }

                WinTitle = "[PS_SD010_02] 생산PO현황";
                ReportName = "PS_SD010_02.rpt";
                //프로시저 : PS_SD010_02

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@ReGbn", ReGbn),
                    new PSH_DataPackClass("@BPLId", BPLId),
                    new PSH_DataPackClass("@ItmBsort", ItmBsort),
                    new PSH_DataPackClass("@CardCode", CardCode),
                    new PSH_DataPackClass("@CntcCode", CntcCode),
                    new PSH_DataPackClass("@RegDateFr", RegDateFr),
                    new PSH_DataPackClass("@RegDateTo", RegDateTo),
                    new PSH_DataPackClass("@RegNumFr", RegNumFr),
                    new PSH_DataPackClass("@RegNumTo", RegNumTo),
                    new PSH_DataPackClass("@ItemCode", ItemCode),
                    new PSH_DataPackClass("@Status", Status)
                };

                string sBPLID = string.Empty;

                if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
                {
                    sBPLID = "전체";
                }
                else
                {
                    sQry01 = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
                    oRecordSet.DoQuery(sQry01);
                    sBPLID = oRecordSet.Fields.Item(0).Value;
                }

                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@BPLId", sBPLID),
                    new PSH_DataPackClass("@RegDateFr", dataHelpClass.ConvertDateType(oForm.Items.Item("RegDateFr").Specific.Value, "-")),
                    new PSH_DataPackClass("@RegDateTo", dataHelpClass.ConvertDateType(oForm.Items.Item("RegDateTo").Specific.Value, "-"))
                };

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_SD010_DeleteHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_SD010_DeleteMatrixSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_SD010_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_SD010_AddMatrixRow(0, true);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD010_DeleteMatrixSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_SD010_AddData() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PS_SD010_LoadCaption();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        if (PS_SD010_DeleteHeaderSpaceLine() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PS_SD010_LoadCaption();
                        PS_SD010_LoadData();
                    }
                    else if (pVal.ItemUID == "Btn03")
                    {
                        if (PS_SD010_DeleteHeaderSpaceLine() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }

                        System.Threading.Thread thread = new System.Threading.Thread(PS_SD010_Print_Report01);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "CardCode")
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
                            if (pVal.ColUID == "ItemCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim()))
                                    {
                                        PSH_Globals.SBO_Application.StatusBar.SetText("사업장 또는 품목대분류를 먼저 선택하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                        return;
                                    }
                                    else
                                    {
                                        PS_SM010 tempForm = new PS_SM010();
                                        tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                        BubbleEvent = false;
                                    }
                                }
                            }
                            else if (pVal.ColUID == "SjDocLin")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("SjDocLin").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
                                    {
                                        PSH_Globals.SBO_Application.StatusBar.SetText("사업장 또는 품목대분류 또는 고객코드를 먼저 선택하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                        return;
                                    }
                                    else
                                    {
                                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                        BubbleEvent = false;
                                    }
                                }
                            }
                            else if (pVal.ColUID == "CntcCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("CntcCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "WhsCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "BItemCD") //원소재코드
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("BItemCD").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "WCntcCD") //생산담당사번
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("WCntcCD").Cells.Item(pVal.Row).Specific.Value))
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
                    if (pVal.ItemUID == "BPLId" || pVal.ItemUID == "ItmBSort")
                    {
                        oForm.Freeze(true);
                        oMat01.Clear();
                        oDS_PS_SD010H.Clear();
                        oMat01.FlushToDataSource();
                        PS_SD010_AddMatrixRow(0, true);
                        oForm.Freeze(false);
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PS_SD010_LoadCaption();
                        }
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
                    if (pVal.ItemUID == "PRType01" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oForm.Freeze(true);
                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();

                        PS_SD010_AddMatrixRow(0, true);
                        oMat01.Columns.Item("SjDocLin").Editable = false;
                        oMat01.Columns.Item("ItemCode").Editable = true;
                        oForm.Freeze(false);
                    }
                    else if (pVal.ItemUID == "PRType02" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oForm.Freeze(true);
                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();

                        PS_SD010_AddMatrixRow(0, true);
                        oMat01.Columns.Item("SjDocLin").Editable = true;
                        oMat01.Columns.Item("ItemCode").Editable = false;
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CntcCode")
                        {
                            PS_SD010_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "CardCode")
                        {
                            PS_SD010_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            PS_SD010_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "SjDocLin" || pVal.ColUID == "ItemCode" || pVal.ColUID == "ReQty" || pVal.ColUID == "CntcCode" || pVal.ColUID == "WhsCode" || pVal.ColUID == "BItemCD" || pVal.ColUID == "WCntcCD")
                            {
                                PS_SD010_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                            }
                            else
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                PS_SD010_LoadCaption();
                            }
                            oMat01.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD010H);
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
            string errMessage = string.Empty;
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
                            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                switch (PSH_Globals.SBO_Application.MessageBox("해당 라인의 생산요청을 삭제합니다. 삭제 후 복원할 수 없습니다. 삭제하시겠습니까?", 1, "&확인", "&취소"))
                                {
                                    case 1: 
                                        //생산의뢰접수여부 체크
                                        sQry = "  SELECT  COUNT(*) AS [Cnt]";
                                        sQry += " FROM    [@PS_PP010H]";
                                        sQry += " WHERE   U_RegNum = '" + oLast_RightClick_RegNum + "'";
                                        sQry += "         AND [Status] = 'O'";
                                        oRecordSet.DoQuery(sQry);

                                        //생산의뢰접수가 존재하면
                                        if (Convert.ToInt32(oRecordSet.Fields.Item("Cnt").Value) > 0)
                                        {
                                            errMessage = "생산의뢰접수가 등록된 작번입니다. 삭제할 수 없습니다.";
                                            BubbleEvent = false;
                                            throw new Exception();
                                        }

                                        sQry = "Delete [@PS_SD010H] Where DocEntry = '" + oLast_RightClick_RegNum + "'";
                                        oRecordSet.DoQuery(sQry);

                                        oLast_RightClick_RegNum = 0;
                                        break;
                                    case 2:
                                        PSH_Globals.SBO_Application.MessageBox("실행이 취소되었습니다.");
                                        BubbleEvent = false;
                                        break;
                                }
                            }
                            break;
                        case "1299": //행닫기
                            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                //생산의뢰접수여부 체크
                                sQry = "  SELECT  COUNT(*) AS [Cnt]";
                                sQry += " FROM    [@PS_PP010H]";
                                sQry += " WHERE   U_RegNum = '" + oLast_RightClick_RegNum + "'";
                                sQry += "         AND [Status] = 'O'";
                                oRecordSet.DoQuery(sQry);

                                if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
                                {
                                    switch (PSH_Globals.SBO_Application.MessageBox("해당 라인의 생산요청을 닫기(취소)합니다. 진행하시겠습니까?", 1, "&확인", "&취소"))
                                    {
                                        case 1:
                                            sQry = "  UPDATE  [@PS_SD010H] ";
                                            sQry += " SET     Status = 'C',";
                                            sQry += "         Canceled = 'Y',";
                                            sQry += "         UpdateDate = GETDATE(),";
                                            sQry += "         UserSign = '" + PSH_Globals.oCompany.UserSignature + "'";
                                            sQry += " Where   U_RegNum = '" + oLast_RightClick_RegNum + "'";
                                            oRecordSet.DoQuery(sQry);

                                            oLast_RightClick_RegNum = 0;
                                            PSH_Globals.SBO_Application.StatusBar.SetText("해당 생산요청건이 닫기(취소)되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                            break;
                                        case 2:
                                            PSH_Globals.SBO_Application.MessageBox("실행이 취소되었습니다.");
                                            BubbleEvent = false;
                                            break;
                                    }
                                }
                                else
                                {
                                    errMessage = "생산의뢰접수된 생산요청입니다. 닫기(취소)할 수 없습니다.";
                                    BubbleEvent = false;
                                    throw new Exception();
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
                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("DocNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }

                                oMat01.FlushToDataSource();
                                oDS_PS_SD010H.RemoveRecord(oDS_PS_SD010H.Size - 1);
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();
                                oForm.Freeze(false);
                            }
                            break;
                        case "1281": //찾기
                            PS_SD010_EnableFormItem();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PS_SD010_LoadCaption();
                            break;
                        case "1282": //추가
                            oForm.Freeze(true);
                            if (oForm.DataSources.UserDataSources.Item("RadioBtn01").Value == "10")
                            {
                                oMat01.Columns.Item("SjDocLin").Editable = false;
                                oMat01.Columns.Item("ItemCode").Editable = true;
                            }
                            else if (oForm.DataSources.UserDataSources.Item("RadioBtn01").Value == "20")
                            {
                                oMat01.Columns.Item("SjDocLin").Editable = true;
                                oMat01.Columns.Item("ItemCode").Editable = false;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_SD010_AddMatrixRow(0, true);
                            PS_SD010_EnableFormItem();
                            PS_SD010_LoadCaption();
                            oForm.Freeze(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_SD010_EnableFormItem();
                            if (oMat01.VisualRowCount > 0)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("CGNo").Cells.Item(oMat01.VisualRowCount).Specific.Value))
                                {
                                    if (oDS_PS_SD010H.GetValue("Status", 0) == "O")
                                    {
                                        PS_SD010_AddMatrixRow(oMat01.RowCount, false);
                                    }
                                }
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
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
                    if (pVal.Row > 0 && oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oLast_RightClick_RegNum = Convert.ToInt32(oMat01.Columns.Item("RegNum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
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

