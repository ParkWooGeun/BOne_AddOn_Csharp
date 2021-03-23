using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 품목별공수조회
	/// </summary>
	internal class PS_PP031 : PSH_BaseClass
	{
		public string oFormUniqueID;
		public SAPbouiCOM.Matrix oMat01;
		public SAPbouiCOM.Matrix oMat02;
		public SAPbouiCOM.Matrix oMat03;
		public SAPbouiCOM.Matrix oMat04;
		private SAPbouiCOM.DBDataSource oDS_PS_PP031H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP031L; //라인(품목분류별규격정보)
		private SAPbouiCOM.DBDataSource oDS_PS_PP031M; //라인(작업지시정보)
		private SAPbouiCOM.DBDataSource oDS_PS_PP031N; //라인(실동공수정보)
		private SAPbouiCOM.DBDataSource oDS_PS_PP031O; //라인(작업일보정보)
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP031.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP031_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP031");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				
				oForm.Freeze(true);
                PS_PP031_CreateItems();
                PS_PP031_ComboBox_Setting();
                PS_PP031_FormResize();

                oForm.Items.Item("Folder01").Specific.Select();
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
        private void PS_PP031_CreateItems()
        {
            try
            {
                oDS_PS_PP031L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_PP031M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
                oDS_PS_PP031N = oForm.DataSources.DBDataSources.Item("@PS_USERDS03");
                oDS_PS_PP031O = oForm.DataSources.DBDataSources.Item("@PS_USERDS04");

                //매트릭스 초기화
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();

                oMat03 = oForm.Items.Item("Mat03").Specific;
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.AutoResizeColumns();

                oMat04 = oForm.Items.Item("Mat04").Specific;
                oMat04.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat04.AutoResizeColumns();

                //품목대분류
                oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");
                
                //품목구분
                oForm.DataSources.UserDataSources.Add("ItemClass", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemClass").Specific.DataBind.SetBound(true, "", "ItemClass");
                
                //거래처구분
                oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");
                
                //품명
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");
                
                //규격_S
                oForm.DataSources.UserDataSources.Add("Spec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Spec").Specific.DataBind.SetBound(true, "", "Spec");
                
                oMat03.Columns.Item("ItemCode").Visible = false; //실동공수정보 Matrix의 품목코드 필드 Hidden
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP031_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //품목대분류
                oForm.Items.Item("ItmBsort").Specific.ValidValues.Add("", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN ('105','106') order by Code", "", false, false);
                oForm.Items.Item("ItmBsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품목구분
                oForm.Items.Item("ItemClass").Specific.ValidValues.Add("", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemClass").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", false, false);
                oForm.Items.Item("ItemClass").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //거래처구분
                oForm.Items.Item("CardType").Specific.ValidValues.Add("", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", false, false);
                oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_PP031_FormResize()
        {
            try
            {
                oForm.Freeze(true);

                oForm.Items.Item("Mat01").Width = oForm.Width / 2 - 20;
                oForm.Items.Item("Mat01").Height = oForm.Height - oForm.Items.Item("Mat01").Top - (oForm.Height - oForm.Items.Item("2").Top) - 10;

                oForm.Items.Item("Mat02").Left = oForm.Width / 2 + 10;
                oForm.Items.Item("Mat02").Width = oForm.Width / 2 - 40;
                oForm.Items.Item("Mat02").Height = oForm.Items.Item("Mat01").Height - 20;

                oForm.Items.Item("Mat03").Left = oForm.Width / 2 + 10;
                oForm.Items.Item("Mat03").Width = oForm.Width / 2 - 40;
                oForm.Items.Item("Mat03").Height = oForm.Items.Item("Mat01").Height - 20;

                oForm.Items.Item("Mat04").Left = oForm.Width / 2 + 10;
                oForm.Items.Item("Mat04").Width = oForm.Width / 2 - 40;
                oForm.Items.Item("Mat04").Height = oForm.Items.Item("Mat01").Height - 20;

                oForm.Items.Item("Rec01").Left = oForm.Width / 2;
                oForm.Items.Item("Rec01").Height = oForm.Items.Item("Mat02").Height + 20;
                oForm.Items.Item("Rec01").Width = oForm.Items.Item("Mat02").Width + 20;

                oForm.Items.Item("Folder01").Width = 100;
                oForm.Items.Item("Folder01").Left = oForm.Items.Item("Rec01").Left;

                oForm.Items.Item("Folder02").Width = 100;
                oForm.Items.Item("Folder02").Left = oForm.Items.Item("Folder01").Left + oForm.Items.Item("Folder02").Width;

                oForm.Items.Item("Folder03").Width = 100;
                oForm.Items.Item("Folder03").Left = oForm.Items.Item("Folder02").Left + oForm.Items.Item("Folder03").Width;

                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();
                oMat03.AutoResizeColumns();
                oMat04.AutoResizeColumns();
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
        /// 메트릭스 Row추가(Mat01)
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param> 
        private void PS_PP031_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_PP031L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP031L.Offset = oRow;
                oMat01.LoadFromDataSource();
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
        /// 메트릭스 데이터 로드
        /// </summary>
        private void PS_PP031_MTX01()
        {
            
            int loopCount;
            string Query01;
            string ItmBsort; //품목대분류
            string CardType; //거래처구분
            string ItemName; //품명
            string ItemClass; //품목구분
            string SPEC; //규격
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);

                ItmBsort = oForm.Items.Item("ItmBsort").Specific.Selected.Value;
                CardType = oForm.Items.Item("CardType").Specific.Selected.Value.ToString().Trim();
                ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToString().Trim();
                ItemName = oForm.Items.Item("ItemName").Specific.Value;
                SPEC = oForm.Items.Item("Spec").Specific.Value;

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                Query01 = "EXEC PS_PP031_01 '" + ItmBsort + "','" + CardType + "','" + ItemClass + "','" + ItemName + "','" + SPEC + "'";
                RecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP031L.InsertRecord(loopCount);
                    }
                    oDS_PS_PP031L.Offset = loopCount;

                    oDS_PS_PP031L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
                    oDS_PS_PP031L.SetValue("U_ColReg01", loopCount, RecordSet01.Fields.Item("ItmBName").Value); //품목대분류명
                    oDS_PS_PP031L.SetValue("U_ColReg02", loopCount, RecordSet01.Fields.Item("ItmMName").Value); //품목구분
                    oDS_PS_PP031L.SetValue("U_ColReg03", loopCount, RecordSet01.Fields.Item("ItemName").Value); //품목명
                    oDS_PS_PP031L.SetValue("U_ColReg04", loopCount, RecordSet01.Fields.Item("Spec").Value); //규격
                    oDS_PS_PP031L.SetValue("U_ColReg05", loopCount, RecordSet01.Fields.Item("Cnt").Value); //Count

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }


        /// <summary>
        /// 메트릭스에 데이터 로드
        /// </summary>
        /// <param name="prmItemName"></param>
        /// <param name="prmSpec"></param>
        private void PS_PP031_MTX02(string prmItemName, string prmSpec)
        {
            int loopCount = 0;
            string Query01 = null;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

                Query01 = "EXEC PS_PP031_02 '" + prmItemName + "','" + prmSpec + "'";
                RecordSet01.DoQuery(Query01);

                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    oMat02.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP031M.InsertRecord(loopCount);
                    }
                    oDS_PS_PP031M.Offset = loopCount;

                    oDS_PS_PP031M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
                    oDS_PS_PP031M.SetValue("U_ColReg01", loopCount, RecordSet01.Fields.Item("WOEntry").Value); //작업지시문서번호
                    oDS_PS_PP031M.SetValue("U_ColReg02", loopCount, RecordSet01.Fields.Item("ItemCode").Value); //품목코드
                    oDS_PS_PP031M.SetValue("U_ColReg03", loopCount, RecordSet01.Fields.Item("ItemName").Value); //품목명
                    oDS_PS_PP031M.SetValue("U_ColReg04", loopCount, RecordSet01.Fields.Item("Spec").Value); //규격
                    oDS_PS_PP031M.SetValue("U_ColReg05", loopCount, RecordSet01.Fields.Item("WODate").Value); //등록일자

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
            }
            catch(Exception ex)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	switch (pVal.EventType) {
        //		case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //			////1
        //			Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //			////2
        //			Raise_EVENT_KEY_DOWN(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //			////5
        //			Raise_EVENT_COMBO_SELECT(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CLICK:
        //			////6
        //			Raise_EVENT_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //			////7
        //			Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //			////8
        //			Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //			////10
        //			Raise_EVENT_VALIDATE(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //			////11
        //			Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //			////18
        //			break;
        //		////et_FORM_ACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //			////19
        //			break;
        //		////et_FORM_DEACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //			////20
        //			Raise_EVENT_RESIZE(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //			////27
        //			Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //			////3
        //			Raise_EVENT_GOT_FOCUS(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //			////4
        //			break;
        //		////et_LOST_FOCUS
        //		case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //			////17
        //			Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //	}
        //	return;
        //	Raise_ItemEvent_Error:
        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				break;
        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;

        //			case "7169":
        //				//엑셀 내보내기

        //				//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
        //				PS_PP031_AddMatrixRow(oMat01.VisualRowCount);
        //				break;

        //		}
        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				break;
        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;

        //			case "7169":
        //				//엑셀 내보내기

        //				//엑셀 내보내기 이후 처리
        //				oForm.Freeze(true);
        //				oDS_PS_PP031L.RemoveRecord(oDS_PS_PP031L.Size - 1);
        //				oMat01.LoadFromDataSource();
        //				oForm.Freeze(false);
        //				break;

        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_FormDataEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //	} else if (pVal.BeforeAction == false) {
        //	}
        //	if (pVal.ItemUID == "Mat01") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	object PP030 = null;
        //	if (pVal.BeforeAction == true) {

        //		if (pVal.ItemUID == "BtnSearch") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (PS_PP031_DataValidCheck() == false) {
        //					BubbleEvent = false;
        //					return;
        //				} else {

        //					//실동공수 정보 초기화
        //					oMat02.Clear();
        //					oDS_PS_PP031M.Clear();

        //					PS_PP031_MTX01();
        //					//매트릭스에 데이터 로드

        //				}
        //			}
        //			//            If oForm.Mode = fm_ADD_MODE Then
        //			//                Call PS_PP031_Print_Report01
        //			//            ElseIf oForm.Mode = fm_UPDATE_MODE Then
        //			//            ElseIf oForm.Mode = fm_OK_MODE Then
        //			//            End If
        //		} else if (pVal.ItemUID == "Btn_Print") {

        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {

        //				PS_PP031_Print_Report01();

        //			}

        //		} else if (pVal.ItemUID == "Link04") {

        //			PP030 = new PS_PP030();

        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: PP030.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			PP030.LoadForm(oForm.Items.Item("WODocNum").Specific.Value);

        //			//UPGRADE_NOTE: PP030 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			PP030 = null;

        //		}
        //	} else if (pVal.BeforeAction == false) {

        //		//폴더를 사용할 때는 필수 소스_S
        //		//Folder01이 선택되었을 때
        //		if (pVal.ItemUID == "Folder01") {

        //			oForm.PaneLevel = 1;

        //		}

        //		//Folder02가 선택되었을 때
        //		if (pVal.ItemUID == "Folder02") {

        //			oForm.PaneLevel = 2;

        //		}

        //		//Folder03이 선택되었을 때
        //		if (pVal.ItemUID == "Folder03") {

        //			oForm.PaneLevel = 3;

        //		}
        //		//폴더를 사용할 때는 필수 소스_E

        //		if (pVal.ItemUID == "PS_PP031") {

        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}

        //		}
        //	}
        //	return;
        //	Raise_EVENT_ITEM_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
        //		//거래처코드 포맷서치 활성
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
        //		//품목코드(작번) 포맷서치 활성

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_KEY_DOWN_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_COMBO_SELECT_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.Row > 0) {

        //				oMat01.SelectRow(pVal.Row, true, false);

        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = pVal.ColUID;
        //				oLastColRow01 = pVal.Row;
        //			}

        //		} else if (pVal.ItemUID == "Mat02") {
        //			if (pVal.Row > 0) {

        //				oMat02.SelectRow(pVal.Row, true, false);

        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = pVal.ColUID;
        //				oLastColRow01 = pVal.Row;
        //			}
        //		} else if (pVal.ItemUID == "Mat03") {
        //			if (pVal.Row > 0) {

        //				oMat03.SelectRow(pVal.Row, true, false);

        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = pVal.ColUID;
        //				oLastColRow01 = pVal.Row;
        //			}
        //		} else if (pVal.ItemUID == "Mat04") {
        //			if (pVal.Row > 0) {

        //				oMat04.SelectRow(pVal.Row, true, false);

        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = pVal.ColUID;
        //				oLastColRow01 = pVal.Row;
        //			}
        //		} else {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = "";
        //			oLastColRow01 = 0;
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //		//품목대분류별 규격 정보
        //		if (pVal.ItemUID == "Mat01") {

        //			if (pVal.Row == 0) {

        //				//정렬
        //				oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
        //				oMat01.FlushToDataSource();

        //			} else {

        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("Folder01").Specific.Select();
        //				//작업지시정보 TAB 선택
        //				//UPGRADE_WARNING: oMat01.Columns(Spec).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				PS_PP031_MTX02(oMat01.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value, oMat01.Columns.Item("Spec").Cells.Item(pVal.Row).Specific.Value);

        //			}

        //		//작업지시정보
        //		} else if (pVal.ItemUID == "Mat02") {

        //			if (pVal.Row == 0) {

        //				//정렬
        //				oMat02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
        //				oMat02.FlushToDataSource();

        //			} else {

        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("Folder02").Specific.Select();
        //				//실동공수정보 TAB 선택
        //				//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				PS_PP031_MTX03(oMat02.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value);

        //			}

        //		//실동공수정보
        //		} else if (pVal.ItemUID == "Mat03") {

        //			if (pVal.Row == 0) {

        //				//정렬
        //				oMat03.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
        //				oMat03.FlushToDataSource();

        //			} else {

        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("Folder03").Specific.Select();
        //				//실동공수정보 TAB 선택
        //				//UPGRADE_WARNING: oMat03.Columns(CpCode).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				PS_PP031_MTX04(oMat03.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value, oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);

        //			}

        //		//작업일보정보
        //		} else if (pVal.ItemUID == "Mat04") {

        //		}

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_DOUBLE_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	object PP030 = null;
        //	object PP040 = null;
        //	if (pVal.BeforeAction == true) {

        //		if (pVal.ItemUID == "Mat02") {

        //			PP030 = new PS_PP030();

        //			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: PP030.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			PP030.LoadForm(oMat02.Columns.Item("WOEntry").Cells.Item(pVal.Row).Specific.Value);

        //			//UPGRADE_NOTE: PP030 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			PP030 = null;

        //		} else if (pVal.ItemUID == "Mat04") {

        //			PP040 = new PS_PP040();

        //			//UPGRADE_WARNING: oMat04.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: PP040.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			PP040.LoadForm(oMat04.Columns.Item("WREntry").Cells.Item(pVal.Row).Specific.Value);

        //			//UPGRADE_NOTE: PP040 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			PP040 = null;

        //		}

        //	} else if (pVal.BeforeAction == false) {

        //	}

        //	return;
        //	Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //	//UPGRADE_NOTE: PP030 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	PP030 = null;
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);

        //	string oQuery01 = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemChanged == true) {

        //			//            If (pVal.ItemUID = "CardCode") Then
        //			//                oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" & oForm.Items(pVal.ItemUID).Specific.Value & "'"
        //			//                oRecordSet01.DoQuery oQuery01
        //			//                oForm.Items("CardName").Specific.Value = Trim(oRecordSet01.Fields(0).Value)
        //			//            ElseIf (pVal.ItemUID = "ItemCode") Then
        //			//                oQuery01 = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm.Items(pVal.ItemUID).Specific.Value & "'"
        //			//                oRecordSet01.DoQuery oQuery01
        //			//                oForm.Items("ItemName").Specific.Value = Trim(oRecordSet01.Fields(0).Value)
        //			//            ElseIf (pVal.ItemUID = "CntcCode") Then
        //			//                oQuery01 = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" & oForm.Items(pVal.ItemUID).Specific.Value & "'"
        //			//                oRecordSet01.DoQuery oQuery01
        //			//                oForm.Items("CntcName").Specific.Value = Trim(oRecordSet01.Fields(0).Value)
        //			//            End If

        //			oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Error:
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		PS_PP031_FormItemEnabled();
        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pVal = null, ref bool BubbleEvent = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		PS_PP031_FormResize();
        //	}
        //	return;
        //	Raise_EVENT_RESIZE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	if (pVal.ItemUID == "Mat01") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else if (pVal.ItemUID == "Mat02") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else if (pVal.ItemUID == "Mat03") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else if (pVal.ItemUID == "Mat04") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}

        //	return;
        //	Raise_EVENT_GOT_FOCUS_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //	} else if (pVal.BeforeAction == false) {
        //		SubMain.RemoveForms(oFormUniqueID01);
        //		//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oForm = null;
        //		//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oMat01 = null;
        //		//UPGRADE_NOTE: oMat02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oMat02 = null;
        //		//UPGRADE_NOTE: oMat03 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oMat03 = null;
        //		//UPGRADE_NOTE: oMat04 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oMat04 = null;
        //	}
        //	return;
        //	Raise_EVENT_FORM_UNLOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion












        #region PS_PP031_MTX03
        //private void PS_PP031_MTX03(string prmItemCode)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////메트릭스에 데이터 로드

        //	int loopCount = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //	ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);
        //	//쿼리를 실행할 때 부터 프로그레스 시작

        //	oForm.Freeze(true);

        //	Query01 = "EXEC PS_PP031_03 '" + prmItemCode + "'";
        //	RecordSet01.DoQuery(Query01);

        //	oMat03.Clear();
        //	oMat03.FlushToDataSource();
        //	oMat03.LoadFromDataSource();

        //	if ((RecordSet01.RecordCount == 0)) {
        //		oMat03.Clear();
        //		goto PS_PP031_MTX03_Exit;
        //	}

        //	for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++) {
        //		if (loopCount != 0) {
        //			oDS_PS_PP031N.InsertRecord(loopCount);
        //		}
        //		oDS_PS_PP031N.Offset = loopCount;

        //		oDS_PS_PP031N.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
        //		//라인번호
        //		oDS_PS_PP031N.SetValue("U_ColReg01", loopCount, RecordSet01.Fields.Item("CpBCode").Value);
        //		//공정대분류
        //		oDS_PS_PP031N.SetValue("U_ColReg02", loopCount, RecordSet01.Fields.Item("CpBName").Value);
        //		//공정대분류명
        //		oDS_PS_PP031N.SetValue("U_ColReg03", loopCount, RecordSet01.Fields.Item("CpCode").Value);
        //		//공정중분류
        //		oDS_PS_PP031N.SetValue("U_ColReg04", loopCount, RecordSet01.Fields.Item("CpName").Value);
        //		//공정중분류명
        //		oDS_PS_PP031N.SetValue("U_ColReg05", loopCount, RecordSet01.Fields.Item("WkTime").Value);
        //		//실동공수
        //		oDS_PS_PP031N.SetValue("U_ColReg06", loopCount, RecordSet01.Fields.Item("ItemCode").Value);
        //		//품목코드

        //		RecordSet01.MoveNext();
        //		ProgressBar01.Value = ProgressBar01.Value + 1;
        //		ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
        //	}
        //	oMat03.LoadFromDataSource();
        //	oMat03.AutoResizeColumns();

        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	return;
        //	PS_PP031_MTX03_Exit:

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	if ((ProgressBar01 != null)) {
        //		ProgressBar01.Stop();
        //	}
        //	MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "W");
        //	return;
        //	PS_PP031_MTX03_Error:
        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP031_MTX03_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP031_MTX04
        //private void PS_PP031_MTX04(string prmItemCode, string prmCpCode)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////메트릭스에 데이터 로드

        //	int loopCount = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //	ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);
        //	//쿼리를 실행할 때 부터 프로그레스 시작

        //	oForm.Freeze(true);

        //	Query01 = "EXEC PS_PP031_04 '" + prmItemCode + "','" + prmCpCode + "'";
        //	RecordSet01.DoQuery(Query01);

        //	oMat04.Clear();
        //	oMat04.FlushToDataSource();
        //	oMat04.LoadFromDataSource();

        //	if ((RecordSet01.RecordCount == 0)) {
        //		oMat04.Clear();
        //		goto PS_PP031_MTX04_Exit;
        //	}

        //	for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++) {
        //		if (loopCount != 0) {
        //			oDS_PS_PP031O.InsertRecord(loopCount);
        //		}
        //		oDS_PS_PP031O.Offset = loopCount;

        //		oDS_PS_PP031O.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
        //		//라인번호
        //		oDS_PS_PP031O.SetValue("U_ColReg01", loopCount, RecordSet01.Fields.Item("WREntry").Value);
        //		//작업일보문서번호
        //		oDS_PS_PP031O.SetValue("U_ColReg02", loopCount, RecordSet01.Fields.Item("ItemCode").Value);
        //		//품목코드
        //		oDS_PS_PP031O.SetValue("U_ColReg03", loopCount, RecordSet01.Fields.Item("ItemName").Value);
        //		//품목명
        //		oDS_PS_PP031O.SetValue("U_ColReg04", loopCount, RecordSet01.Fields.Item("Spec").Value);
        //		//규격
        //		oDS_PS_PP031O.SetValue("U_ColReg05", loopCount, RecordSet01.Fields.Item("CpCode").Value);
        //		//공정코드
        //		oDS_PS_PP031O.SetValue("U_ColReg06", loopCount, RecordSet01.Fields.Item("CpName").Value);
        //		//공정명
        //		oDS_PS_PP031O.SetValue("U_ColReg07", loopCount, RecordSet01.Fields.Item("WkTime").Value);
        //		//실동공수
        //		oDS_PS_PP031O.SetValue("U_ColReg08", loopCount, RecordSet01.Fields.Item("WRDate").Value);
        //		//등록일자

        //		RecordSet01.MoveNext();
        //		ProgressBar01.Value = ProgressBar01.Value + 1;
        //		ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
        //	}
        //	oMat04.LoadFromDataSource();
        //	oMat04.AutoResizeColumns();

        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	return;
        //	PS_PP031_MTX04_Exit:

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	if ((ProgressBar01 != null)) {
        //		ProgressBar01.Stop();
        //	}
        //	MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "W");
        //	return;
        //	PS_PP031_MTX04_Error:
        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP031_MTX04_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion


        #region PS_PP031_Print_Report01
        //private void PS_PP031_Print_Report01()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocNum = null;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;

        //	short i = 0;
        //	short ErrNum = 0;
        //	string Sub_sQry = null;

        //	//    Dim BPLId           As String
        //	//    Dim CardCode       As String
        //	//    Dim Pumtxt        As String


        //	SAPbobsCOM.Recordset oRecordSet = null;

        //	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	MDC_PS_Common.ConnectODBC();

        //	//// 조회조건문

        //	//    BPLId = Trim(oForm.Items("BPLId").Specific.Value)
        //	//    CardCode = Trim(oForm.Items("CardCode").Specific.Value)
        //	//    Pumtxt = Trim(oForm.Items("Pumtxt").Specific.Value)
        //	//
        //	//    If Pumtxt = "" Then Pumtxt = "%"
        //	//    If CardCode = "" Then CardCode = "%"
        //	//

        //	string BPLID = null;
        //	//사업장
        //	string ItemClass = null;
        //	//품목구분
        //	string TradeType = null;
        //	//거래형태
        //	string FrDt = null;
        //	//납기일시작
        //	string ToDt = null;
        //	//납기일종료
        //	string CardCode = null;
        //	//거래처
        //	string ItemCode = null;
        //	//품목코드(작번)
        //	string DocStatus = null;
        //	//문서상태
        //	string Chk01 = null;
        //	//미출고
        //	string Chk02 = null;
        //	//미납품

        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	BPLID = Strings.Trim(oForm.Items.Item("BPLId").Specific.Selected.Value);
        //	//UPGRADE_WARNING: oForm.Items(ItemClass).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ItemClass = (oForm.Items.Item("ItemClass").Specific.Selected.Value == "%" ? "" : oForm.Items.Item("ItemClass").Specific.Selected.Value);
        //	//UPGRADE_WARNING: oForm.Items(TradeType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	TradeType = (oForm.Items.Item("TradeType").Specific.Selected.Value == "%" ? "" : oForm.Items.Item("TradeType").Specific.Selected.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	FrDt = Strings.Trim(oForm.Items.Item("FrDt").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ToDt = Strings.Trim(oForm.Items.Item("ToDt").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	ItemCode = Strings.Trim(oForm.Items.Item("ItemCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items(DocStatus).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocStatus = (oForm.Items.Item("DocStatus").Specific.Selected.Value == "%" ? "" : oForm.Items.Item("DocStatus").Specific.Selected.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Chk01 = (oForm.Items.Item("Chk01").Specific.Checked == true ? "1" : "0");
        //	//UPGRADE_WARNING: oForm.Items().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Chk02 = (oForm.Items.Item("Chk02").Specific.Checked == true ? "1" : "0");


        //	/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        //	WinTitle = "[PS_PP031] 레포트";
        //	ReportName = "PS_PP031.rpt";
        //	MDC_Globals.gRpt_Formula = new string[3];
        //	MDC_Globals.gRpt_Formula_Value = new string[3];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //	//// Formula 수식필드

        //	//// SubReport


        //	MDC_Globals.gRpt_SFormula[1, 1] = "";
        //	MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

        //	/// Procedure 실행"
        //	sQry = "EXEC PS_PP031_01 '" + BPLID + "','" + ItemClass + "','" + TradeType + "','" + FrDt + "','" + ToDt + "','" + CardCode + "','" + ItemCode + "','" + DocStatus + "','" + Chk01 + "','" + Chk02 + "'";

        //	oRecordSet.DoQuery(sQry);
        //	if (oRecordSet.RecordCount == 0) {
        //		ErrNum = 1;
        //		goto Print_Query_Error;
        //	}

        //	/// Action (sub_query가 있을때는 'Y'로...)/
        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") == false) {
        //	}

        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;
        //	return;
        //	Print_Query_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet = null;
        //	if (ErrNum == 1) {
        //		MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
        //	} else {
        //		MDC_Com.MDC_GF_Message(ref "Print_Query_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //	}
        //}
        #endregion
    }
}
