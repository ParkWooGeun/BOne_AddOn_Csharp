using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 결산마감관리
	/// </summary>
	internal class PS_CO001 : PSH_BaseClass
	{
		private string oFormUniqueID;
		//public SAPbouiCOM.Form oForm01;
		private SAPbouiCOM.Matrix oMat01;
	
		private SAPbouiCOM.DBDataSource oDS_PS_CO001L; //라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO001.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO001_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO001");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
                PS_CO001_CreateItems();
                PS_CO001_ComboBox_Setting();
                PS_CO001_Initial_Setting();
                PS_CO001_SetDocument(oFromDocEntry01);
                PS_CO001_FormResize();
                PS_CO001_LoadCaption();
                PS_CO001_FormItemEnabled();

                oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1285", false); //복원
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1293", false); //행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
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
        private void PS_CO001_CreateItems()
        {
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                oDS_PS_CO001L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                //사업장
                oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");
                
                //기준년도
                oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void PS_CO001_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                //사업장
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

                sQry = "  SELECT    B.U_Minor, ";
                sQry += "           B.U_CdName";
                sQry += " FROM      [@PS_SY001H] A";
                sQry += "           INNER JOIN";
                sQry += "           [@PS_SY001L] B";
                sQry += "               ON A.Code = B.Code";
                sQry += " WHERE     A.Code = 'F006'";

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("MM005Sts"), sQry, "", ""); //구매요청
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("PP040Sts"), sQry, "", ""); //작업일보등록
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("PP080Sts"), sQry, "", ""); //생산완료
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("SD040Sts"), sQry, "", ""); //납품처리
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ARSts"), sQry, "", ""); //AR송장
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("APSts"), sQry, "", ""); //AP송장
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("OIGESts"), sQry, "", ""); //출고
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
        /// 초기 세팅
        /// </summary>
        private void PS_CO001_Initial_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("StdYear").Specific.Value = DateTime.Now.ToString("yyyy"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PS_CO001_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PS_CO001_FormItemEnabled();
                }
                else
                {
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Form Resize
        /// </summary>
        private void PS_CO001_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
        /// </summary>
        private void PS_CO001_LoadCaption()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "저장";

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
                }
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
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_CO001_FormItemEnabled()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ///메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_CO001_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_CO001L.InsertRecord((oRow));
                }

                oMat01.AddRow();
                oDS_PS_CO001L.Offset = oRow;
                oDS_PS_CO001L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 매트릭스 데이터 로드
        /// </summary>
        private void PS_CO001_MTX01()
        {
            short loopCount;
            string sQry;
            string errCode = string.Empty;
            string BPLID; //사업장
            string StdYear; //기준년도

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                BPLID = oForm.Items.Item("BPLID").Specific.Selected.Value;
                StdYear = oForm.Items.Item("StdYear").Specific.Value;

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

                oForm.Freeze(true);

                sQry = "EXEC [PS_CO001_01] '";
                sQry += BPLID + "','";
                sQry += StdYear + "'";

                RecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PS_CO001L.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    errCode = "1";
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PS_CO001_LoadCaption();
                    throw new Exception();
                }

                for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount + 1 > oDS_PS_CO001L.Size)
                    {
                        oDS_PS_CO001L.InsertRecord(loopCount);
                    }

                    oMat01.AddRow();
                    oDS_PS_CO001L.Offset = loopCount;

                    oDS_PS_CO001L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
                    oDS_PS_CO001L.SetValue("U_ColReg01", loopCount, RecordSet01.Fields.Item("StdMonth").Value); //기준월
                    oDS_PS_CO001L.SetValue("U_ColReg02", loopCount, RecordSet01.Fields.Item("MM005Sts").Value); //구매요청
                    oDS_PS_CO001L.SetValue("U_ColReg03", loopCount, RecordSet01.Fields.Item("PP040Sts").Value); //작업일보등록
                    oDS_PS_CO001L.SetValue("U_ColReg04", loopCount, RecordSet01.Fields.Item("PP080Sts").Value); //생산완료
                    oDS_PS_CO001L.SetValue("U_ColReg05", loopCount, RecordSet01.Fields.Item("SD040Sts").Value); //납품처리
                    oDS_PS_CO001L.SetValue("U_ColReg06", loopCount, RecordSet01.Fields.Item("ARSts").Value); //AR송장
                    oDS_PS_CO001L.SetValue("U_ColReg07", loopCount, RecordSet01.Fields.Item("APSts").Value); //AP송장
                    oDS_PS_CO001L.SetValue("U_ColReg08", loopCount, RecordSet01.Fields.Item("OIGESts").Value); //출고
                    oDS_PS_CO001L.SetValue("U_ColReg09", loopCount, RecordSet01.Fields.Item("Comments").Value); //비고

                    RecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Update();
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        /// <summary>
        /// 데이터 INSERT
        /// </summary>
        private void PS_CO001_AddData()
        {
            string sQry;
            string StdYear; //년도
            string BPLID; //사업장
            string StdMonth; //기준월
            string MM005Sts; //구매요청
            string PP040Sts; //작업일보등록
            string PP080Sts; //생산완료
            string SD040Sts; //납품처리
            string ARSts; //AR송장
            string APSts; //AP송장
            string OIGESts; //출고
            string Comments; //비고

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();

                for (loopCount = 0; loopCount <= oMat01.RowCount - 1; loopCount++)
                {
                    BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim(); //사업장
                    StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim(); //년도
                    StdMonth = oDS_PS_CO001L.GetValue("U_ColReg01", loopCount).ToString().Trim(); //기준월
                    MM005Sts = oDS_PS_CO001L.GetValue("U_ColReg02", loopCount).ToString().Trim(); //구매요청
                    PP040Sts = oDS_PS_CO001L.GetValue("U_ColReg03", loopCount).ToString().Trim(); //작업일보등록
                    PP080Sts = oDS_PS_CO001L.GetValue("U_ColReg04", loopCount).ToString().Trim(); //생산완료
                    SD040Sts = oDS_PS_CO001L.GetValue("U_ColReg05", loopCount).ToString().Trim(); //납품처리
                    ARSts = oDS_PS_CO001L.GetValue("U_ColReg06", loopCount).ToString().Trim(); //AR송장
                    APSts = oDS_PS_CO001L.GetValue("U_ColReg07", loopCount).ToString().Trim(); //AP송장
                    OIGESts = oDS_PS_CO001L.GetValue("U_ColReg08", loopCount).ToString().Trim(); //출고
                    Comments = oDS_PS_CO001L.GetValue("U_ColReg09", loopCount).ToString().Trim(); //비고

                    sQry = "EXEC [PS_CO001_02] '";
                    sQry += BPLID + "','"; //사업장
                    sQry += StdYear + "','"; //년도
                    sQry += StdMonth + "','"; //기준월
                    sQry += MM005Sts + "','"; //구매요청
                    sQry += PP040Sts + "','"; //작업일보등록
                    sQry += PP080Sts + "','"; //생산완료
                    sQry += SD040Sts + "','"; //납품처리
                    sQry += ARSts + "','"; //AR송장
                    sQry += APSts + "','"; //AP송장
                    sQry += OIGESts + "','"; //출고
                    sQry += Comments + "'"; //비고

                    RecordSet01.DoQuery(sQry);
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }
        

        #region PS_CO001_HeaderSpaceLineDel
        //		private bool PS_CO001_HeaderSpaceLineDel()
        //		{
        //			bool functionReturnValue = false;
        //			//******************************************************************************
        //			//Function ID : PS_CO001_HeaderSpaceLineDel()
        //			//해당모듈    : PS_CO001
        //			//기능        : 필수입력사항 체크
        //			//인수        : 없음
        //			//반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
        //			//특이사항    : 없음
        //			//******************************************************************************
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			short ErrNum = 0;
        //			ErrNum = 0;

        //			//UPGRADE_WARNING: oForm01.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			switch (true) {
        //				case Strings.Trim(oForm01.Items.Item("BPLID").Specific.VALUE) == "%":
        //					//사업장
        //					ErrNum = 1;
        //					goto PS_CO001_HeaderSpaceLineDel_Error;
        //					break;
        //				case string.IsNullOrEmpty(Strings.Trim(oForm01.Items.Item("StdYear").Specific.VALUE)):
        //					//기준년도
        //					ErrNum = 2;
        //					goto PS_CO001_HeaderSpaceLineDel_Error;
        //					break;
        //			}

        //			functionReturnValue = true;
        //			return functionReturnValue;
        //			PS_CO001_HeaderSpaceLineDel_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			if (ErrNum == 1) {
        //				MDC_Com.MDC_GF_Message(ref "사업장을 선택하지 않았습니다. 확인하세요.", ref "E");
        //				oForm01.Items.Item("BPLID").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else if (ErrNum == 2) {
        //				MDC_Com.MDC_GF_Message(ref "기준년도을 입력하지 않았습니다.. 확인하세요.", ref "E");
        //				oForm01.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			}
        //			functionReturnValue = false;
        //			return functionReturnValue;
        //		}
        #endregion

        #region PS_CO001_FlushToItemValue
        //		private void PS_CO001_FlushToItemValue(string oUID, ref int oRow = 0, ref string oCol = "")
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			short i = 0;
        //			short ErrNum = 0;
        //			string sQry = null;
        //			string ItemCode = null;
        //			short Qty = 0;
        //			string Total = null;

        //			SAPbobsCOM.Recordset oRecordSet01 = null;
        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			switch (oUID) {

        //				case "Mat01":

        //					oMat01.FlushToDataSource();
        //					break;

        //			}

        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;

        //			return;
        //			PS_CO001_FlushToItemValue_Error:

        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			oRecordSet01 = null;

        //			if (ErrNum == 1) {
        //				MDC_Com.MDC_GF_Message(ref "에러 메시지", ref "E");
        //			} else {
        //				MDC_Com.MDC_GF_Message(ref "PS_CO001_FlushToItemValue_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			}
        //		}
        #endregion



        #region PS_CO001_FormReset
        //		public void PS_CO001_FormReset()
        //		{
        //			//******************************************************************************
        //			//Function ID : PS_CO001_FormReset()
        //			//해당모듈    : PS_CO001
        //			//기능        : 화면 초기화
        //			//인수        : 없음
        //			//반환값      : 없음
        //			//특이사항    : 없음
        //			//******************************************************************************
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			string sQry = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			oForm01.Freeze(true);

        //			//    '관리번호
        //			//    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_CO001H]"
        //			//    Call RecordSet01.DoQuery(sQry)
        //			//
        //			//    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        //			//        Call oDS_PS_CO001H.setValue("DocEntry", 0, 1)
        //			//    Else
        //			//        Call oDS_PS_CO001H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
        //			//    End If

        //			//    Call oDS_PS_CO001H.setValue("DocEntry", 0, "") '관리번호
        //			//    Call oDS_PS_CO001H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId) '사업장
        //			//    Call oDS_PS_CO001H.setValue("U_ZGODAY", 0, Format(Date, "YYYYMMDD")) '기준일자
        //			//    Call oDS_PS_CO001H.setValue("U_ZEMPIE", 0, "") '관련자사번
        //			//    Call oDS_PS_CO001H.setValue("U_ZEMPNM", 0, "") '관련자성명
        //			//    Call PS_CO001_FlushToItemValue("BPLID") '관련자사업장에 따른 팀 세팅
        //			//    Call oDS_PS_CO001H.setValue("U_ZDPTCD", 0, "%") '관련자팀
        //			//    Call PS_CO001_FlushToItemValue("ZDPTCD") '관련자팀에 따른 담당 세팅
        //			//    Call oDS_PS_CO001H.setValue("U_ZSECCD", 0, "%") '관련자담당
        //			//    Call oDS_PS_CO001H.setValue("U_ZPOINT", 0, "%") '장소
        //			//    Call oDS_PS_CO001H.setValue("U_ZUSRNN", 0, "") '방문자성명
        //			//    Call oDS_PS_CO001H.setValue("U_ZBIRDT", 0, "") '생년월일
        //			//    Call oDS_PS_CO001H.setValue("U_ZBUSNM", 0, "") '업체명
        //			//    Call oDS_PS_CO001H.setValue("U_ZRESID", 0, "") '주소
        //			//    Call oDS_PS_CO001H.setValue("U_ZGOTIME", 0, "") '출입시간(Fr)
        //			//    Call oDS_PS_CO001H.setValue("U_ZOUTIME", 0, "") '출입시간(To)
        //			//    Call oDS_PS_CO001H.setValue("U_ZCARJO", 0, "%") '차량종류
        //			//    Call oDS_PS_CO001H.setValue("U_ZCARNO", 0, "") '차량번호
        //			//    Call oDS_PS_CO001H.setValue("U_ZMOKJU", 0, "%") '출입목적
        //			//    Call oDS_PS_CO001H.setValue("U_ZSPCIU", 0, "") '방문객비고

        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			oForm01.Freeze(false);

        //			return;
        //			PS_CO001_FormReset_Error:

        //			oForm01.Freeze(false);
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			MDC_Com.MDC_GF_Message(ref "PS_CO001_FormReset_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //		}
        #endregion







        #region Raise_ItemEvent
        /////아이템 변경 이벤트
        //		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			switch (pval.EventType) {
        //				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //					////1
        //					Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //					////2
        //					Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //					////5
        //					Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CLICK:
        //					////6
        //					Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //					////7
        //					Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //					////8
        //					Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //					////10
        //					Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //					////11
        //					Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //					////18
        //					break;
        //				////et_FORM_ACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //					////19
        //					break;
        //				////et_FORM_DEACTIVATE
        //				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //					////20
        //					Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //					////27
        //					Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //					////3
        //					Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //					////4
        //					break;
        //				////et_LOST_FOCUS
        //				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //					////17
        //					Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
        //					break;
        //			}
        //			return;
        //			Raise_ItemEvent_Error:
        //			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_MenuEvent
        //		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			string sQry = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1293":
        //						//행삭제
        //						Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
        //						break;
        //					case "1281":
        //						//찾기
        //						break;
        //					case "1282":
        //						//추가
        //						///추가버튼 클릭시 메트릭스 insertrow

        //						//                Call PS_CO001_FormReset

        //						//                oMat01.Clear
        //						//                oMat01.FlushToDataSource
        //						//                oMat01.LoadFromDataSource

        //						oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //						BubbleEvent = false;
        //						PS_CO001_LoadCaption();

        //						//oForm01.Items("GCode").Click ct_Regular


        //						return;

        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;

        //					case "7169":
        //						//엑셀 내보내기

        //						//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
        //						PS_CO001_Add_MatrixRow(oMat01.VisualRowCount);
        //						break;


        //				}
        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1293":
        //						//행삭제
        //						Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
        //						break;
        //					case "1281":
        //						//찾기
        //						break;
        //					////Call PS_CO001_FormItemEnabled '//UDO방식
        //					case "1282":
        //						//추가
        //						break;
        //					//                oMat01.Clear
        //					//                oDS_PS_CO001H.Clear

        //					//                Call PS_CO001_LoadCaption
        //					//                Call PS_CO001_FormItemEnabled
        //					////Call PS_CO001_FormItemEnabled '//UDO방식
        //					////Call PS_CO001_AddMatrixRow(0, True) '//UDO방식
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //					////Call PS_CO001_FormItemEnabled

        //					case "7169":
        //						//엑셀 내보내기

        //						//엑셀 내보내기 이후 처리
        //						oForm01.Freeze(true);
        //						oDS_PS_CO001L.RemoveRecord(oDS_PS_CO001L.Size - 1);
        //						oMat01.LoadFromDataSource();
        //						oForm01.Freeze(false);
        //						break;

        //				}
        //			}
        //			return;
        //			Raise_MenuEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((BusinessObjectInfo.BeforeAction == true)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			////BeforeAction = False
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //				switch (BusinessObjectInfo.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //						////33
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //						////34
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //						////35
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //						////36
        //						break;
        //				}
        //			}
        //			return;
        //			Raise_FormDataEvent_Error:
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //			}

        //			if (pval.ItemUID == "Mat01") {
        //				if (pval.Row > 0) {
        //					oLastItemUID01 = pval.ItemUID;
        //					oLastColUID01 = pval.ColUID;
        //					oLastColRow01 = pval.Row;
        //				}
        //			} else {
        //				oLastItemUID01 = pval.ItemUID;
        //				oLastColUID01 = "";
        //				oLastColRow01 = 0;
        //			}

        //			return;
        //			Raise_RightClickEvent_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //				if (pval.ItemUID == "PS_CO001") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}

        //				//추가(수정) 버튼클릭
        //				if (pval.ItemUID == "BtnAdd") {

        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {

        //						if (PS_CO001_HeaderSpaceLineDel() == false) {
        //							BubbleEvent = false;
        //							return;
        //						}

        //						PS_CO001_AddData();

        //						//                Call PS_CO001_FormReset
        //						oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

        //						PS_CO001_LoadCaption();
        //						PS_CO001_MTX01();

        //						//oLast_Mode = oForm01.Mode

        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {

        //						//                If PS_CO001_HeaderSpaceLineDel = False Then
        //						//                    BubbleEvent = False
        //						//                    Exit Sub
        //						//                End If

        //						//쿼리에서 로직 처리
        //						//여러 행을 동시에 처리하므로 DocEntry별로 행이 존재하면 UPDATE, 존재하지 않으면 INSERT
        //						PS_CO001_AddData();

        //						//                Call PS_CO001_FormReset
        //						oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

        //						PS_CO001_LoadCaption();
        //						PS_CO001_MTX01();

        //						//                oForm01.Items("GCode").Click ct_Regular
        //					}

        //				} else if (pval.ItemUID == "BtnSrch") {

        //					//            Call PS_CO001_FormReset
        //					oForm01.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //					///fm_VIEW_MODE

        //					if (PS_CO001_HeaderSpaceLineDel() == false) {
        //						BubbleEvent = false;
        //						return;
        //					}

        //					PS_CO001_LoadCaption();
        //					PS_CO001_MTX01();

        //				}

        //			} else if (pval.BeforeAction == false) {
        //				if (pval.ItemUID == "PS_CO001") {
        //					if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //					} else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					}
        //				}
        //			}

        //			return;
        //			Raise_EVENT_ITEM_PRESSED_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "AcctCode01", "");
        //				////사용자값활성
        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "AcctCode02", "");
        //				////사용자값활성
        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "AcctCode03", "");
        //				////사용자값활성

        //				//        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_KEY_DOWN_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_CLICK
        //		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //				if (pval.ItemUID == "Mat01") {

        //					if (pval.Row > 0) {

        //						oMat01.SelectRow(pval.Row, true, false);

        //					}

        //				}

        //			}

        //			return;
        //			Raise_EVENT_CLICK_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //				PS_CO001_FlushToItemValue(pval.ItemUID);

        //			}

        //			return;
        //			Raise_EVENT_COMBO_SELECT_Error:

        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_DOUBLE_CLICK_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_MATRIX_LINK_PRESSED_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_VALIDATE
        //		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm01.Freeze(true);

        //			if (pval.BeforeAction == true) {

        //				if (pval.ItemChanged == true) {

        //					if (pval.ItemUID == "Mat01") {

        //						if (pval.ColUID == "Month01" | pval.ColUID == "Month02" | pval.ColUID == "Month03" | pval.ColUID == "Month04" | pval.ColUID == "Month05" | pval.ColUID == "Month06" | pval.ColUID == "Month07" | pval.ColUID == "Month08" | pval.ColUID == "Month09" | pval.ColUID == "Month10" | pval.ColUID == "Month11" | pval.ColUID == "Month12") {

        //							PS_CO001_FlushToItemValue(pval.ItemUID, ref pval.Row, ref pval.ColUID);

        //						}

        //					} else {

        //						PS_CO001_FlushToItemValue(pval.ItemUID);

        //					}

        //				}

        //				//            oMat01.LoadFromDataSource
        //				//            oMat01.AutoResizeColumns
        //				//            oForm01.Update

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			oForm01.Freeze(false);

        //			return;
        //			Raise_EVENT_VALIDATE_Error:

        //			oForm01.Freeze(false);
        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {
        //				PS_CO001_FormItemEnabled();
        //				////Call PS_CO001_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        //			}

        //			return;
        //			Raise_EVENT_MATRIX_LOAD_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_RESIZE
        //		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //				PS_CO001_FormResize();

        //			}

        //			return;
        //			Raise_EVENT_RESIZE_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {
        //				//        If (pval.ItemUID = "ItemCode") Then
        //				//            Dim oDataTable01 As SAPbouiCOM.DataTable
        //				//            Set oDataTable01 = pval.SelectedObjects
        //				//            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
        //				//            Set oDataTable01 = Nothing
        //				//        End If
        //				//        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
        //				//            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_CO001H", "U_CardCode,U_CardName")
        //				//        End If
        //			}

        //			return;
        //			Raise_EVENT_CHOOSE_FROM_LIST_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.ItemUID == "Mat01") {
        //				if (pval.Row > 0) {
        //					oLastItemUID01 = pval.ItemUID;
        //					oLastColUID01 = pval.ColUID;
        //					oLastColRow01 = pval.Row;
        //				}
        //			} else {
        //				oLastItemUID01 = pval.ItemUID;
        //				oLastColUID01 = "";
        //				oLastColRow01 = 0;
        //			}

        //			return;
        //			Raise_EVENT_GOT_FOCUS_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {
        //				SubMain.RemoveForms(oFormUniqueID01);
        //				//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oForm01 = null;
        //				//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				oMat01 = null;
        //			}

        //			return;
        //			Raise_EVENT_FORM_UNLOAD_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			int i = 0;

        //			if ((oLastColRow01 > 0)) {
        //				if (pval.BeforeAction == true) {
        //					//            If (PS_CO001_Validate("행삭제") = False) Then
        //					//                BubbleEvent = False
        //					//                Exit Sub
        //					//            End If
        //					////행삭제전 행삭제가능여부검사

        //				} else if (pval.BeforeAction == false) {
        //					for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
        //					}
        //					oMat01.FlushToDataSource();
        //					oDS_PS_CO001L.RemoveRecord(oDS_PS_CO001L.Size - 1);
        //					oMat01.LoadFromDataSource();

        //					//            If oMat01.RowCount = 0 Then
        //					//                Call PS_CO001_Add_MatrixRow(0)
        //					//            Else
        //					//                If Trim(oDS_PS_CO001L.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
        //					//                    Call PS_CO001_Add_MatrixRow(oMat01.RowCount)
        //					//                End If
        //					//            End If

        //				}
        //			}

        //			return;
        //			Raise_EVENT_ROW_DELETE_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion


    }
}
