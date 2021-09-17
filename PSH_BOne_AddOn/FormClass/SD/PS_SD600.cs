using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 견적서등록
	/// </summary>
	internal class PS_SD600 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SD600H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD600L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD600.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD600_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD600");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_SD600_CreateItems();
                PS_SD600_ComboBox_Setting();
                PS_SD600_Initial_Setting();
                PS_SD600_EnableMenus();
                PS_SD600_SetDocument(oFormDocEntry);
                //PS_SD600_FormResize();

                oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1286", true); //닫기
				oForm.EnableMenu("1284", true); //취소
				oForm.EnableMenu("1293", true); //행삭제
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
        private void PS_SD600_CreateItems()
        {
            try
            {
                oDS_PS_SD600H = oForm.DataSources.DBDataSources.Item("@PS_SD600H");
                oDS_PS_SD600L = oForm.DataSources.DBDataSources.Item("@PS_SD600L");
                oMat01 = oForm.Items.Item("Mat01").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD600_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {   //사업부
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                //을병지포함
                oForm.Items.Item("SubYN").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SubYN").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'M001' ORDER BY Code", "", false, false);
                oForm.Items.Item("SubYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품목구분
                oForm.Items.Item("ItemGrp").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemGrp").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", false, false);
                oForm.Items.Item("ItemGrp").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //단위
                oForm.Items.Item("Unit").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("Unit").Specific, "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code", "", false, false);
                oForm.Items.Item("Unit").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //매트릭스-구분_S
                sQry = "SELECT  U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S004' ORDER BY U_Minor";
                oRecordSet01.DoQuery(sQry);

                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("Class").ValidValues.Add(oRecordSet01.Fields.Item(0).Value, oRecordSet01.Fields.Item(1).Value);
                    oRecordSet01.MoveNext();
                }
                //매트릭스-구분_E

                //매트릭스-단위_S
                sQry = "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code";
                oRecordSet01.DoQuery(sQry);

                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("Unit").ValidValues.Add(oRecordSet01.Fields.Item(0).Value, oRecordSet01.Fields.Item(1).Value);
                    oRecordSet01.MoveNext();
                }
                //매트릭스-단위_E
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
        /// 초기 세팅
        /// </summary>
        private void PS_SD600_Initial_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //사업장
                oForm.Items.Item("QuotDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //견적일자
                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD(); //견적등록자
                oDS_PS_SD600H.SetValue("U_PayCon", 0, "귀사 정기 지불"); //지불조건
                oDS_PS_SD600H.SetValue("U_DueDate", 0, "발주 후 60일"); //납기일
                oDS_PS_SD600H.SetValue("U_DelCon", 0, "귀사 지정"); //물품인도조건
                oDS_PS_SD600H.SetValue("U_QuotDue", 0, "30일"); //견적유효기간
                oForm.Items.Item("SubYN").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue); //을병지포함 선택
                oForm.Items.Item("ItemGrp").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue); //품목구분 선택
                oForm.Items.Item("Unit").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue); //단위 선택
                oForm.Items.Item("CardName").Click(SAPbouiCOM.BoCellClickType.ct_Regular); //최초 거래처 텍스트박스 선택
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_SD600_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
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
        private void PS_SD600_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_SD600_FormItemEnabled();
                    PS_SD600_AddMatrixRow(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_SD600_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param> 
        private void PS_SD600_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_SD600L.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PS_SD600L.Offset = oRow;
                oDS_PS_SD600L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_SD600_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PS_SD600_FormItemEnabled
        /// </summary>
        private void PS_SD600_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("QuotNo").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
                    PS_SD600_FormClear();
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "";
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("QuotNo").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("QuotNo").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_SD600_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD600'", "");

                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
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
        /// 필수입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_SD600_DataValidCheck()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            int i = 0;

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_SD600_FormClear();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("QuotDt").Specific.Value)) //견적일자
                {
                    errCode = "1";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value)) //등록자
                {
                    errCode = "2";
                    throw new Exception();
                }
                else if (oForm.Items.Item("SubYN").Specific.Value == "%") //을병지포함
                {
                    errCode = "3";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("CardName").Specific.Value)) //거래처명
                {
                    errCode = "4";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("CardPart").Specific.Value)) //부서
                {
                    errCode = "5";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("CardPrsn").Specific.Value)) //담당자
                {
                    errCode = "6";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("QuotNm").Specific.Value)) //견적명
                {
                    errCode = "7";
                    throw new Exception();
                }
                else if (oForm.Items.Item("SubYN").Specific.Value == "Y") //품명(을병지를 포함할 때만 품명 체크)
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("ItemName").Specific.Value))
                    {
                        errCode = "8";
                        throw new Exception();
                    }
                }
                else if (oForm.Items.Item("SubYN").Specific.Value == "Y") //도번,규격(을병지를 포함할 때만 품명 체크)
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("Spec").Specific.Value))
                    {
                        errCode = "9";
                        throw new Exception();
                    }
                }
                else if (oForm.Items.Item("ItemGrp").Specific.Value == "%") //품목구분
                {
                    errCode = "10";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1) //라인정보 미입력 시
                {
                    errCode = "11";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("Class").Cells.Item(i).Specific.Value)) //구분
                    {
                        errCode = "12";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("Content1").Cells.Item(i).Specific.Value)) //내용(재질/규격/공정)
                    {
                        errCode = "13";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("Content2").Cells.Item(i).Specific.Value)) //수량(중량/공수)
                    {
                        errCode = "14";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("Unit").Cells.Item(i).Specific.Value)) //단위
                    {
                        errCode = "15";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value)) //단가
                    {
                        errCode = "16";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat01.Columns.Item("Amount").Cells.Item(i).Specific.Value)) //금액
                    {
                        errCode = "17";
                        throw new Exception();
                    }
                }

                oMat01.FlushToDataSource();
                oDS_PS_SD600L.RemoveRecord(oDS_PS_SD600L.Size - 1);
                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("견적일자가 입력되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("QuotDt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("등록자가 입력되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("을병지포함이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("SubYN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("거래처명이 입력되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CardName").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "5")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("거래처부서가 입력되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CardPart").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "6")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("거래처담당자가 입력되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CardPrsn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "7")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("견적명이 입력되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("QuotNm").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "8")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("품명이 입력되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("ItemName").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "9")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("도번/규격이 입력되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("Spec").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "10")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("품목구분이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("ItemGrp").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "11")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인이 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "12")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("구분은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("Class").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "13")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("내용(재질/규격/공정)은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("Content1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "14")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("수량(중량/공수)은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("Content2").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "15")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("검진구분은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("Unit").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "16")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("단가는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("Price").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "17")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("금액은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("Amount").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 견적번호 자동생성
        /// </summary>
        /// <param name="prmBPLId">사업장</param>
        /// <param name="prmQuotDt">견적일자</param>
        /// <returns>견적번호</returns>
        private string PS_SD600_MakeQutationNo(string prmBPLId, string prmQuotDt)
        {
            string sQry;
            string returnValue = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "EXEC PS_SD600_90 '";
                sQry += prmBPLId + "','"; //사업장
                sQry += prmQuotDt + "'"; //견적일자
                oRecordSet01.DoQuery(sQry);

                returnValue = oRecordSet01.Fields.Item("QuotNo").Value;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 견적서 리포트 출력
        /// </summary>
        [STAThread]
        private void PS_SD600_Print_Report01()
        {
            string WinTitle;
            string ReportName;

            string quotNo; //견적번호
            string subYN; //을병지 포함여부
            
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                quotNo = oForm.Items.Item("QuotNo").Specific.Value.ToString().Trim();
                subYN = oForm.Items.Item("SubYN").Specific.Value.ToString().Trim();

                WinTitle = "견적서 갑지";
                ReportName = "PS_SD600_51.rpt";
                //프로시저 : PS_SD600_01

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                
                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@QuotNo", quotNo)); //견적번호
                dataPackParameter.Add(new PSH_DataPackClass("@SubYM", subYN)); //을병지 포함여부
                
                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
            string QuotNo;
            string BPLID;
            string QuotDt;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_SD600_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                            QuotDt = oForm.Items.Item("QuotDt").Specific.Value.ToString().Trim();
                            QuotNo = PS_SD600_MakeQutationNo(BPLID, QuotDt); //견적번호 자동 생성

                            if (QuotNo == "Overflow") //견적번호 오버플로우
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                oDS_PS_SD600H.SetValue("U_QuotNo", 0, QuotNo);
                            }

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD600_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "btnPrint1")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_SD600_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_SD600_FormItemEnabled();
                            }
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
                        if (pVal.ColUID == "Content1") //내용(재질/규격/공정)
                        {
                            if (pVal.CharPressed == 9) //탭을 눌렀을 경우만
                            {
                                dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "Content1"); //내용(재질/규격/공정) 포맷서치설정
                            }
                        }
                    }
                    else if (pVal.ItemUID == "CntcCode")
                    {
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", ""); //견적등록자 포맷서치 설정
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
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "Class")
                            {
                                oMat01.FlushToDataSource();

                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_SD600L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_SD600_AddMatrixRow(pVal.Row, false);
                                }
                            }
                        }
                    }
                }
                oMat01.AutoResizeColumns();
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
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;

                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
    
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {   
                            if (pVal.ColUID == "Content1") //발생공정코드
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_SD600L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oMat01.LoadFromDataSource();

                            }
                            else if (pVal.ColUID == "Content2" || pVal.ColUID == "Price")
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_SD600L.SetValue("U_Amount", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oDS_PS_SD600L.GetValue("U_Content2", pVal.Row - 1)) * Convert.ToDouble(oDS_PS_SD600L.GetValue("U_Price", pVal.Row - 1))));
                                oMat01.LoadFromDataSource();
                            }

                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oMat01.AutoResizeColumns();
                        }
                        else if (pVal.ItemUID == "CntcCode")
                        {
                            oDS_PS_SD600H.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
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
                oForm.Freeze(false);
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
                    PS_SD600_FormItemEnabled();
                    PS_SD600_AddMatrixRow(oMat01.VisualRowCount, false);
                    oMat01.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD600H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD600L);
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
                    PS_SD600_FormResize();
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
        /// ROW_DELETE 이벤트
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.MenuEvent pVal, bool BubbleEvent)
        { 
            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (int i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_SD600L.RemoveRecord(oDS_PS_SD600L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_SD600_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_SD600L.GetValue("U_Class", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_SD600_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
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
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_SD600_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PS_SD600_FormItemEnabled(); 
                            PS_SD600_AddMatrixRow(0, true);
                            PS_SD600_Initial_Setting();
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            PS_SD600_FormItemEnabled();
                            break;
                        case "1287":
                            oForm.Freeze(true);
                            oDS_PS_SD600H.SetValue("DocEntry", 0, "");

                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_SD600H.SetValue("DocEntry", i, "");
                                oMat01.LoadFromDataSource();
                            }

                            oForm.Freeze(false);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
