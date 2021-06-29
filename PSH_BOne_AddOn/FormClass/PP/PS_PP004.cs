using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 품목별 표준 공정 등록
	/// </summary>
	internal class PS_PP004 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.Matrix oMat03;
		private SAPbouiCOM.DBDataSource oDS_PS_USERDS01; //품목
		private SAPbouiCOM.DBDataSource oDS_PS_USERDS02; //공정
		private SAPbouiCOM.DBDataSource oDS_PS_USERDS03; //공정추가
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oMat01Row01;
		private int oMat02Row02;
		private int oMat03Row03;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP004.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP004_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP004");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_PP004_CreateItems();
				PS_PP004_SetComboBox();
				PS_PP004_CF_ChooseFromList();
				PS_PP004_EnableMenus();
				PS_PP004_ResizeForm();
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
        private void PS_PP004_CreateItems()
        {
            try
            {
                oDS_PS_USERDS01 = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_USERDS02 = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
                oDS_PS_USERDS03 = oForm.DataSources.DBDataSources.Item("@PS_USERDS03");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat03 = oForm.Items.Item("Mat03").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();
                oMat03.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

                oForm.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");

                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");

                oForm.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");

                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt03");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP004_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_Insert("PS_PP004", "Mat03", "ResultYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP004", "Mat03", "ResultYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("ResultYN"), "PS_PP004", "Mat03", "ResultYN", false);
                dataHelpClass.Combo_ValidValues_Insert("PS_PP004", "Mat03", "ReportYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP004", "Mat03", "ReportYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("ReportYN"), "PS_PP004", "Mat03", "ReportYN", false);
                dataHelpClass.Combo_ValidValues_Insert("PS_PP004", "Mat03", "DayProYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP004", "Mat03", "DayProYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("DayProYN"), "PS_PP004", "Mat03", "DayProYN", false);
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", false, false);

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmsGrpCod"), "SELECT ItmsGrpCod, ItmsGrpNam FROM OITB order by ItmsGrpCod", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmMsort"), "SELECT Code, Name FROM [@PSH_ITMMSORT] order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Mark"), "SELECT Code, Name FROM [@PSH_MARK] order by Code", "", "");

                dataHelpClass.GP_MatrixSetMatComboList(oMat02.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code", "", "");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList 설정
        /// </summary>
        private void PS_PP004_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.EditText oEdit = null;

            try
            {
                oEdit = oForm.Items.Item("ItemCode").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLITEMCODE";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oEdit.ChooseFromListUID = "CFLITEMCODE";
                oEdit.ChooseFromListAlias = "ItemCode";
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP004_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, false, false, false, false, false, false, false, false, false, false, false, false, false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ResizeForm
        /// </summary>
        private void PS_PP004_ResizeForm()
        {
            try
            {
                oForm.Items.Item("Mat01").Top = 90;
                oForm.Items.Item("Mat01").Height = (oForm.ClientHeight) - 120;
                oForm.Items.Item("Mat01").Left = 7;
                oForm.Items.Item("Mat01").Width = oForm.ClientWidth / 2 - 10;

                oForm.Items.Item("Mat02").Top = 90;
                oForm.Items.Item("Mat02").Height = ((oForm.ClientHeight - 120) / 2) - 15;
                oForm.Items.Item("Mat02").Left = oForm.ClientWidth / 2;
                oForm.Items.Item("Mat02").Width = oForm.ClientWidth / 2 - 10;

                oForm.Items.Item("Mat03").Top = 90 + ((oForm.ClientHeight - 120) / 2) + 15;
                oForm.Items.Item("Mat03").Height = ((oForm.ClientHeight - 120) / 2) - 15;
                oForm.Items.Item("Mat03").Left = oForm.ClientWidth / 2;
                oForm.Items.Item("Mat03").Width = oForm.ClientWidth / 2 - 10;

                oForm.Items.Item("Button02").Top = oForm.Items.Item("Mat03").Top - 25;
                oForm.Items.Item("Button03").Top = oForm.Items.Item("Mat03").Top - 25;
                oForm.Items.Item("Button04").Top = oForm.Items.Item("Mat03").Top - 25;

                oForm.Items.Item("Button02").Left = (oForm.Items.Item("Mat03").Left + oForm.Items.Item("Mat03").Width) - 100;
                oForm.Items.Item("Button03").Left = (oForm.Items.Item("Mat03").Left + oForm.Items.Item("Mat03").Width) - 202;
                oForm.Items.Item("Button04").Left = (oForm.Items.Item("Mat03").Left + oForm.Items.Item("Mat03").Width) - 304;

                oForm.Items.Item("Opt03").Top = oForm.Items.Item("Mat03").Top - 20;
                oForm.Items.Item("Opt03").Left = oForm.Items.Item("Mat03").Left;

                oForm.Items.Item("Opt02").Left = oForm.Items.Item("Mat02").Left;

                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();
                oMat03.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP004_CheckDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (oForm.Items.Item("ItmBsort").Specific.Selected == null)
                {
                    errMessage = "대분류는 필수입니다.";
                    oForm.Items.Item("ItmBsort").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        /// 메트릭스 데이터 로드(품목정보)
        /// </summary>
        private void PS_PP004_MTX01()
        {
            int i;
            string Query01;
            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03;
            string errMessage = string.Empty;

            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                if (oForm.Items.Item("ItmBsort").Specific.Selected != null)
                {
                    Param01 = oForm.Items.Item("ItmBsort").Specific.Selected.Value.ToString().Trim(); //대분류
                }

                if (oForm.Items.Item("ItmMsort").Specific.Selected != null)
                {
                    Param02 = oForm.Items.Item("ItmMsort").Specific.Selected.Value.ToString().Trim(); //중분류
                }

                Param03 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim(); //품목코드

                Query01 = "EXEC PS_PP004_01 '" + Param01 + "','" + Param02 + "','" + Param03 + "'";
                RecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                oMat03.Clear();
                oMat03.FlushToDataSource();
                oMat03.LoadFromDataSource();

                oMat01Row01 = 0;
                oMat02Row02 = 0;
                oMat03Row03 = 0;

                if (RecordSet01.RecordCount == 0)
                {
                    errMessage = "품목정보 결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                //멀티,엔드베어링
                if (RecordSet01.Fields.Item("ItmBsort").Value == "107" || RecordSet01.Fields.Item("ItmBsort").Value == "104")
                {
                    oForm.Items.Item("Button03").Enabled = false;
                    oForm.Items.Item("Button04").Enabled = false;
                    oMat03.Columns.Item("ResultYN").Editable = false;
                    oMat03.Columns.Item("ReportYN").Editable = false;
                    oMat03.Columns.Item("DayProYN").Editable = false;
                }
                else
                {
                    oForm.Items.Item("Button03").Enabled = true;
                    oForm.Items.Item("Button04").Enabled = true;
                    oMat03.Columns.Item("ResultYN").Editable = true;
                    oMat03.Columns.Item("ReportYN").Editable = true;
                    oMat03.Columns.Item("DayProYN").Editable = true;
                }

                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_USERDS01.InsertRecord(i);
                    }
                    oDS_PS_USERDS01.Offset = i;
                    oDS_PS_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_USERDS01.SetValue("U_ColReg01", i, RecordSet01.Fields.Item("ItemCode").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg02", i, RecordSet01.Fields.Item("ItemName").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg03", i, RecordSet01.Fields.Item("ItmsGrpCod").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg04", i, RecordSet01.Fields.Item("ItmBsort").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg05", i, RecordSet01.Fields.Item("ItmMsort").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg06", i, RecordSet01.Fields.Item("Unit1").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg07", i, RecordSet01.Fields.Item("Size").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg08", i, RecordSet01.Fields.Item("ItemType").Value);
                    oDS_PS_USERDS01.SetValue("U_ColQty01", i, RecordSet01.Fields.Item("UnWeight").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg09", i, RecordSet01.Fields.Item("Quality").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg10", i, RecordSet01.Fields.Item("Mark").Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg11", i, RecordSet01.Fields.Item("CallSize").Value);
                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();                
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
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// 메트릭스 데이터 로드(공정정보)
        /// </summary>
        private void PS_PP004_MTX02()
        {
            int i;
            string Query01;
            string Param01 = string.Empty;
            string errMessage = string.Empty;

            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                if (oForm.Items.Item("ItmBsort").Specific.Selected != null)
                {
                    Param01 = oForm.Items.Item("ItmBsort").Specific.Selected.Value.ToString().Trim(); //대분류
                }

                Query01 = "EXEC PS_PP004_02 '" + Param01 + "'";
                RecordSet01.DoQuery(Query01);

                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    errMessage = "공정정보 결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_USERDS02.InsertRecord(i);
                    }
                    oDS_PS_USERDS02.Offset = i;
                    oDS_PS_USERDS02.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_USERDS02.SetValue("U_ColReg01", i, RecordSet01.Fields.Item("CpBCode").Value);
                    oDS_PS_USERDS02.SetValue("U_ColReg02", i, RecordSet01.Fields.Item("CpBName").Value);
                    oDS_PS_USERDS02.SetValue("U_ColReg03", i, RecordSet01.Fields.Item("CpCode").Value);
                    oDS_PS_USERDS02.SetValue("U_ColReg04", i, RecordSet01.Fields.Item("CpName").Value);
                    oDS_PS_USERDS02.SetValue("U_ColReg05", i, RecordSet01.Fields.Item("ItmBsort").Value);
                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
                oForm.Update();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// 메트릭스 데이터 로드(공정추가)
        /// </summary>
        private void PS_PP004_MTX03()
        {
            int i;
            string Query01;
            string Param01;
            string errMessage = string.Empty;

            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                Param01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value;

                Query01 = "EXEC PS_PP004_03 '" + Param01 + "'";
                RecordSet01.DoQuery(Query01);

                oMat03.Clear();
                oMat03.FlushToDataSource();
                oMat03.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    errMessage = "공정추가 결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_USERDS03.InsertRecord(i);
                    }
                    oDS_PS_USERDS03.Offset = i;
                    oDS_PS_USERDS03.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_USERDS03.SetValue("U_ColNum01", i, RecordSet01.Fields.Item("Sequence").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg01", i, RecordSet01.Fields.Item("ItemCode").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg02", i, RecordSet01.Fields.Item("ItemName").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg03", i, RecordSet01.Fields.Item("CpBCode").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg04", i, RecordSet01.Fields.Item("CpBName").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg05", i, RecordSet01.Fields.Item("CpCode").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg06", i, RecordSet01.Fields.Item("CpName").Value);
                    oDS_PS_USERDS03.SetValue("U_ColQty01", i, RecordSet01.Fields.Item("CpUnWt").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg07", i, RecordSet01.Fields.Item("ResultYN").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg08", i, RecordSet01.Fields.Item("ReportYN").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg10", i, RecordSet01.Fields.Item("DayProYN").Value);
                    oDS_PS_USERDS03.SetValue("U_ColReg09", i, RecordSet01.Fields.Item("StdTime").Value);
                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat03.LoadFromDataSource();
                oMat03.AutoResizeColumns();
                oForm.Update();
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
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// RegisteProcess
        /// </summary>
        /// <returns></returns>
        private bool PS_PP004_RegisteProcess()
        {
            bool returnValue = false;
            int i;
            int Code;
            int Count;
            bool CP30112 = false;
            bool CP30114 = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                for (i = 1; i <= oMat03.VisualRowCount; i++)
                {
                    if (oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.Value == "CP301") //휘팅공정
                    {
                        if (oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.Value == "CP30112") //바렐공정
                        {
                            CP30112 = true;
                        }
                        else if (oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.Value == "CP30114") //포장공정
                        {
                            CP30114 = true;
                        }
                    }
                }

                for (i = 1; i <= oMat03.VisualRowCount; i++)
                {
                    if (oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.Value == "CP301") //휘팅공정
                    {
                        if (CP30112 != true || CP30114 != true)
                        {
                            errMessage = "휘팅공정은 바렐,포장공정은 필수 입니다.";
                            throw new Exception();
                        }
                    }
                }

                Count = 0;

                for (i = 1; i <= oMat03.VisualRowCount; i++)
                {
                    if (oMat03.Columns.Item("DayProYN").Cells.Item(i).Specific.Value == "Y") //일생산실적
                    {
                        Count += 1;
                    }
                }

                if (Count > 1)
                {
                    errMessage = "일생산실적 포인트는 하나만 가능합니다. 저장되지 않았습니다.";
                    throw new Exception();
                }

                dataHelpClass.DoQuery("DELETE [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value + "'");

                for (i = 1; i <= oMat03.VisualRowCount; i++)
                {
                    Code = Convert.ToInt16(dataHelpClass.GetValue("SELECT ISNULL(MAX(CONVERT(INT,Code)),0) + 1 FROM [@PS_PP004H]", 0, 1));
                    dataHelpClass.DoQuery("INSERT INTO [@PS_PP004H] VALUES ('" + Code + "','" + Code + "','" + oMat03.Columns.Item("Sequence").Cells.Item(i).Specific.Value + "','" + oMat03.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + "','" + oMat03.Columns.Item("ItemName").Cells.Item(i).Specific.Value + "','" + oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.Value + "','" + oMat03.Columns.Item("CpBName").Cells.Item(i).Specific.Value + "','" + oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.Value + "','" + oMat03.Columns.Item("CpName").Cells.Item(i).Specific.Value + "','" + oMat03.Columns.Item("CpUnWt").Cells.Item(i).Specific.Value + "','" + oMat03.Columns.Item("StdTime").Cells.Item(i).Specific.Value + "','" + oMat03.Columns.Item("ResultYN").Cells.Item(i).Specific.Selected.Value + "','" + oMat03.Columns.Item("ReportYN").Cells.Item(i).Specific.Selected.Value + "','" + oMat03.Columns.Item("DayProYN").Cells.Item(i).Specific.Selected.Value + "')");
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("데이터가 " + oMat03.VisualRowCount + " 개 저장되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
            int Sequence;
            string ItemCode;
            string ItemName;
            string CpBCode;
            string CpBName;
            string CpCode;
            string CpName;
            double CpUnWt;
            string ResultYN;
            string ReportYN;
            string DayProYN;
            int CurrentMat03Row;

            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP004_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PS_PP004_MTX01();
                            PS_PP004_MTX02();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Button02")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP004_RegisteProcess() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Button03") //위로
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (oMat03.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) == -1)
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("행이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            }
                            else if (oMat03.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) == 1)
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("첫 행입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            }
                            else
                            {
                                CurrentMat03Row = oMat03Row03;
                                Sequence = Convert.ToInt16(oMat03.Columns.Item("Sequence").Cells.Item(CurrentMat03Row).Specific.Value);
                                ItemCode = oMat03.Columns.Item("ItemCode").Cells.Item(CurrentMat03Row).Specific.Value;
                                ItemName = oMat03.Columns.Item("ItemName").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpCode = oMat03.Columns.Item("CpCode").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpName = oMat03.Columns.Item("CpName").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpBCode = oMat03.Columns.Item("CpBCode").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpBName = oMat03.Columns.Item("CpBName").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpUnWt = Convert.ToDouble(oMat03.Columns.Item("CpUnWt").Cells.Item(CurrentMat03Row).Specific.Value);
                                ResultYN = oMat03.Columns.Item("ResultYN").Cells.Item(CurrentMat03Row).Specific.Selected.Value;
                                ReportYN = oMat03.Columns.Item("ReportYN").Cells.Item(CurrentMat03Row).Specific.Selected.Value;
                                DayProYN = oMat03.Columns.Item("DayProYN").Cells.Item(CurrentMat03Row).Specific.Selected.Value;

                                oMat03.Columns.Item("Sequence").Cells.Item(CurrentMat03Row).Specific.Value = Convert.ToInt16(oMat03.Columns.Item("Sequence").Cells.Item(CurrentMat03Row - 1).Specific.Value) + 1;
                                oMat03.Columns.Item("Sequence").Cells.Item(CurrentMat03Row - 1).Specific.Value = Sequence - 1;
                                oMat03.Columns.Item("ItemCode").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("ItemCode").Cells.Item(CurrentMat03Row - 1).Specific.Value;
                                oMat03.Columns.Item("ItemCode").Cells.Item(CurrentMat03Row - 1).Specific.Value = ItemCode;
                                oMat03.Columns.Item("ItemName").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("ItemName").Cells.Item(CurrentMat03Row - 1).Specific.Value;
                                oMat03.Columns.Item("ItemName").Cells.Item(CurrentMat03Row - 1).Specific.Value = ItemName;
                                oMat03.Columns.Item("CpBCode").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpBCode").Cells.Item(CurrentMat03Row - 1).Specific.Value;
                                oMat03.Columns.Item("CpBCode").Cells.Item(CurrentMat03Row - 1).Specific.Value = CpBCode;
                                oMat03.Columns.Item("CpBName").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpBName").Cells.Item(CurrentMat03Row - 1).Specific.Value;
                                oMat03.Columns.Item("CpBName").Cells.Item(CurrentMat03Row - 1).Specific.Value = CpBName;
                                oMat03.Columns.Item("CpCode").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpCode").Cells.Item(CurrentMat03Row - 1).Specific.Value;
                                oMat03.Columns.Item("CpCode").Cells.Item(CurrentMat03Row - 1).Specific.Value = CpCode;
                                oMat03.Columns.Item("CpName").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpName").Cells.Item(CurrentMat03Row - 1).Specific.Value;
                                oMat03.Columns.Item("CpName").Cells.Item(CurrentMat03Row - 1).Specific.Value = CpName;
                                oMat03.Columns.Item("CpUnWt").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpUnWt").Cells.Item(CurrentMat03Row - 1).Specific.Value;
                                oMat03.Columns.Item("CpUnWt").Cells.Item(CurrentMat03Row - 1).Specific.Value = CpUnWt;
                                oMat03.Columns.Item("ResultYN").Cells.Item(CurrentMat03Row).Specific.Select(oMat03.Columns.Item("ResultYN").Cells.Item(CurrentMat03Row - 1).Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("ResultYN").Cells.Item(CurrentMat03Row - 1).Specific.Select(ResultYN, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("ReportYN").Cells.Item(CurrentMat03Row).Specific.Select(oMat03.Columns.Item("ReportYN").Cells.Item(CurrentMat03Row - 1).Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("ReportYN").Cells.Item(CurrentMat03Row - 1).Specific.Select(ReportYN, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("DayProYN").Cells.Item(CurrentMat03Row).Specific.Select(oMat03.Columns.Item("DayProYN").Cells.Item(CurrentMat03Row - 1).Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("DayProYN").Cells.Item(CurrentMat03Row - 1).Specific.Select(DayProYN, SAPbouiCOM.BoSearchKey.psk_ByValue);

                                oMat03.SelectRow(CurrentMat03Row - 1, true, false);
                                oMat03Row03 = CurrentMat03Row - 1;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Button04") //아래로
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (oMat03.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) == -1)
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("행이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            }
                            else if (oMat03.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) == oMat03.VisualRowCount)
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("마지막 행입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            }
                            else
                            {
                                CurrentMat03Row = oMat03Row03;
                                Sequence = Convert.ToInt16(oMat03.Columns.Item("Sequence").Cells.Item(CurrentMat03Row).Specific.Value);
                                ItemCode = oMat03.Columns.Item("ItemCode").Cells.Item(CurrentMat03Row).Specific.Value;
                                ItemName = oMat03.Columns.Item("ItemName").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpBCode = oMat03.Columns.Item("CpBCode").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpBName = oMat03.Columns.Item("CpBName").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpCode = oMat03.Columns.Item("CpCode").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpName = oMat03.Columns.Item("CpName").Cells.Item(CurrentMat03Row).Specific.Value;
                                CpUnWt = Convert.ToDouble(oMat03.Columns.Item("CpUnWt").Cells.Item(CurrentMat03Row).Specific.Value);
                                ResultYN = oMat03.Columns.Item("ResultYN").Cells.Item(CurrentMat03Row).Specific.Selected.Value;
                                ReportYN = oMat03.Columns.Item("ReportYN").Cells.Item(CurrentMat03Row).Specific.Selected.Value;
                                DayProYN = oMat03.Columns.Item("DayProYN").Cells.Item(CurrentMat03Row).Specific.Selected.Value;

                                oMat03.Columns.Item("Sequence").Cells.Item(CurrentMat03Row).Specific.Value = Convert.ToInt16(oMat03.Columns.Item("Sequence").Cells.Item(CurrentMat03Row + 1).Specific.Value) - 1;
                                oMat03.Columns.Item("Sequence").Cells.Item(CurrentMat03Row + 1).Specific.Value = Sequence + 1;
                                oMat03.Columns.Item("ItemCode").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("ItemCode").Cells.Item(CurrentMat03Row + 1).Specific.Value;
                                oMat03.Columns.Item("ItemCode").Cells.Item(CurrentMat03Row + 1).Specific.Value = ItemCode;
                                oMat03.Columns.Item("ItemName").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("ItemName").Cells.Item(CurrentMat03Row + 1).Specific.Value;
                                oMat03.Columns.Item("ItemName").Cells.Item(CurrentMat03Row + 1).Specific.Value = ItemName;
                                oMat03.Columns.Item("CpBCode").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpBCode").Cells.Item(CurrentMat03Row + 1).Specific.Value;
                                oMat03.Columns.Item("CpBCode").Cells.Item(CurrentMat03Row + 1).Specific.Value = CpBCode;
                                oMat03.Columns.Item("CpBName").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpBName").Cells.Item(CurrentMat03Row + 1).Specific.Value;
                                oMat03.Columns.Item("CpBName").Cells.Item(CurrentMat03Row + 1).Specific.Value = CpBName;
                                oMat03.Columns.Item("CpCode").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpCode").Cells.Item(CurrentMat03Row + 1).Specific.Value;
                                oMat03.Columns.Item("CpCode").Cells.Item(CurrentMat03Row + 1).Specific.Value = CpCode;
                                oMat03.Columns.Item("CpName").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpName").Cells.Item(CurrentMat03Row + 1).Specific.Value;
                                oMat03.Columns.Item("CpName").Cells.Item(CurrentMat03Row + 1).Specific.Value = CpName;
                                oMat03.Columns.Item("CpUnWt").Cells.Item(CurrentMat03Row).Specific.Value = oMat03.Columns.Item("CpUnWt").Cells.Item(CurrentMat03Row + 1).Specific.Value;
                                oMat03.Columns.Item("CpUnWt").Cells.Item(CurrentMat03Row + 1).Specific.Value = CpUnWt;
                                oMat03.Columns.Item("ResultYN").Cells.Item(CurrentMat03Row).Specific.Select(oMat03.Columns.Item("ResultYN").Cells.Item(CurrentMat03Row + 1).Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("ResultYN").Cells.Item(CurrentMat03Row + 1).Specific.Select(ResultYN, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("ReportYN").Cells.Item(CurrentMat03Row).Specific.Select(oMat03.Columns.Item("ReportYN").Cells.Item(CurrentMat03Row + 1).Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("ReportYN").Cells.Item(CurrentMat03Row + 1).Specific.Select(ReportYN, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("DayProYN").Cells.Item(CurrentMat03Row).Specific.Select(oMat03.Columns.Item("DayProYN").Cells.Item(CurrentMat03Row + 1).Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMat03.Columns.Item("DayProYN").Cells.Item(CurrentMat03Row + 1).Specific.Select(DayProYN, SAPbouiCOM.BoSearchKey.psk_ByValue);

                                oMat03.SelectRow(CurrentMat03Row + 1, true, false);
                                oMat03Row03 = CurrentMat03Row + 1;
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                    if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02" || pVal.ItemUID == "Mat03")
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
                        }
                        else
                        {
                            if (pVal.ItemUID == "ItmBsort")
                            {
                                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value + "' order by CONVERT(INT,U_CODE)", "", true, false);
                                oForm.Items.Item("ItmMsort").Specific.ValidValues.Add("", "");
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oForm.Update();
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
                    if (pVal.ItemUID == "Opt01")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat01";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oForm.Freeze(false);
                    }
                    else if (pVal.ItemUID == "Opt02")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat02";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oForm.Freeze(false);
                    }
                    else if (pVal.ItemUID == "Opt03")
                    {
                        oForm.Freeze(true);
                        oForm.Settings.MatrixUID = "Mat03";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oForm.Freeze(false);
                    }

                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                            oMat01Row01 = pVal.Row;
                            PS_PP004_MTX03();
                        }
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat02.SelectRow(pVal.Row, true, false);
                            oMat02Row02 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat03.SelectRow(pVal.Row, true, false);
                            oMat03Row03 = pVal.Row;
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

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            if (oMat01Row01 == 0)
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("품목이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                for (i = 1; i <= oMat03.VisualRowCount; i++)
                                {
                                    //공정매트릭스와 품목별공정결과 매트릭스에 똑같은 공정이 존재할경우
                                    if (oMat02.Columns.Item("CpBCode").Cells.Item(pVal.Row).Specific.Value == oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.Value && oMat02.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value == oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.Value)
                                    {
                                        PSH_Globals.SBO_Application.StatusBar.SetText("품목이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                        BubbleEvent = false;
                                        return;
                                    }
                                }

                                //멀티,엔드베어링
                                if (oMat01.Columns.Item("ItmBsort").Cells.Item(oMat01Row01).Specific.Value == "104" || oMat01.Columns.Item("ItmBsort").Cells.Item(oMat01Row01).Specific.Value == "107")
                                {
                                    for (i = 1; i <= oMat02.VisualRowCount; i++)
                                    {
                                        oMat03.AddRow();
                                        oMat03.Columns.Item("LineNum").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat03.VisualRowCount;
                                        oMat03.Columns.Item("Sequence").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat03.VisualRowCount;
                                        oMat03.Columns.Item("ItemCode").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value;
                                        oMat03.Columns.Item("ItemName").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat01.Columns.Item("ItemName").Cells.Item(oMat01Row01).Specific.Value;
                                        oMat03.Columns.Item("CpBCode").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat02.Columns.Item("CpBCode").Cells.Item(i).Specific.Value;
                                        oMat03.Columns.Item("CpBName").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat02.Columns.Item("CpBName").Cells.Item(i).Specific.Value;
                                        oMat03.Columns.Item("CpCode").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat02.Columns.Item("CpCode").Cells.Item(i).Specific.Value;
                                        oMat03.Columns.Item("CpName").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat02.Columns.Item("CpName").Cells.Item(i).Specific.Value;
                                        oMat03.Columns.Item("CpUnWt").Cells.Item(oMat03.VisualRowCount).Specific.Value = 0;

                                        if (i == oMat02.VisualRowCount && oMat01.Columns.Item("ItmBsort").Cells.Item(oMat01Row01).Specific.Value != "104")
                                        {
                                            oMat03.Columns.Item("ResultYN").Cells.Item(oMat03.VisualRowCount).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }
                                        else if (oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.Value == "CP50107")
                                        {
                                            oMat03.Columns.Item("ResultYN").Cells.Item(oMat03.VisualRowCount).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }
                                        else
                                        {
                                            oMat03.Columns.Item("ResultYN").Cells.Item(oMat03.VisualRowCount).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        }
                                        oMat03.Columns.Item("ReportYN").Cells.Item(oMat03.VisualRowCount).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oMat03.Columns.Item("DayProYN").Cells.Item(oMat03.VisualRowCount).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    }
                                }
                                else
                                {
                                    oMat03.AddRow();
                                    oMat03.Columns.Item("LineNum").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat03.VisualRowCount;
                                    oMat03.Columns.Item("Sequence").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat03.VisualRowCount;
                                    oMat03.Columns.Item("ItemCode").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value;
                                    oMat03.Columns.Item("ItemName").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat01.Columns.Item("ItemName").Cells.Item(oMat01Row01).Specific.Value;
                                    oMat03.Columns.Item("CpBCode").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat02.Columns.Item("CpBCode").Cells.Item(oMat02Row02).Specific.Value;
                                    oMat03.Columns.Item("CpBName").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat02.Columns.Item("CpBName").Cells.Item(oMat02Row02).Specific.Value;
                                    oMat03.Columns.Item("CpCode").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat02.Columns.Item("CpCode").Cells.Item(oMat02Row02).Specific.Value;
                                    oMat03.Columns.Item("CpName").Cells.Item(oMat03.VisualRowCount).Specific.Value = oMat02.Columns.Item("CpName").Cells.Item(oMat02Row02).Specific.Value;
                                    oMat03.Columns.Item("CpUnWt").Cells.Item(oMat03.VisualRowCount).Specific.Value = 0;
                                    oMat03.Columns.Item("ResultYN").Cells.Item(oMat03.VisualRowCount).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oMat03.Columns.Item("ReportYN").Cells.Item(oMat03.VisualRowCount).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oMat03.Columns.Item("DayProYN").Cells.Item(oMat03.VisualRowCount).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
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
                oForm.Freeze(false);
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
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "StdTime")
                            {
                                oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = codeHelpClass.Left(Convert.ToString(Convert.ToInt32(oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)), 4);
                            }
                        }
                        else
                        {
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_USERDS01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_USERDS02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_USERDS03);
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
                    PS_PP004_ResizeForm();
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (oDataTable01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
                }
            }
        }

        /// <summary>
        /// 행삭제 체크 메소드(Raise_FormMenuEvent 에서 사용)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                int i = 0;
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (oMat03.Columns.Item("CpCode").Cells.Item(oMat03Row03).Specific.Value == "CP30112" || oMat03.Columns.Item("CpCode").Cells.Item(oMat03Row03).Specific.Value == "CP30114") //바렐공정이거나 포장공정
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("바렐공정 또는 포장공정은 행삭제 할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (oMat01.Columns.Item("ItmBsort").Cells.Item(oMat01Row01).Specific.Value == "104" || oMat01.Columns.Item("ItmBsort").Cells.Item(oMat01Row01).Specific.Value == "107")
                        {
                        }
                        else
                        {
                            for (i = 1; i <= oMat03.VisualRowCount; i++)
                            {
                                oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                                oMat03.Columns.Item("Sequence").Cells.Item(i).Specific.Value = i;
                            }
                        }
                    }
                }
                else
                {
                    BubbleEvent = false;
                    return;
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "Mat03")
                    {
                        oForm.EnableMenu("1293", true);
                    }
                    else
                    {
                        oForm.EnableMenu("1293", false);
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }

                if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02" || pVal.ItemUID == "Mat03")
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
