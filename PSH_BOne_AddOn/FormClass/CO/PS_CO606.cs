using System;

using SAPbouiCOM;
using SAP.Middleware.Connector;

using PSH_BOne_AddOn.Data;
using SAPbobsCOM;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 통합재무제표 본사 전송
	/// </summary>
	internal class PS_CO606 : PSH_BaseClass
	{
		public string oFormUniqueID;
		//public SAPbouiCOM.Form oForm01;
		public SAPbouiCOM.Matrix oMat01;
		private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLast_Col_Row; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO606.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO606_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO606");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
				//oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				CreateItems();
				
				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
        private void CreateItems()
        {
            try
            {
                //기준년월F
                oForm.DataSources.UserDataSources.Add("StdYMF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("StdYMF").Specific.DataBind.SetBound(true, "", "StdYMF");
                oForm.DataSources.UserDataSources.Item("StdYMF").Value = DateTime.Now.ToString("yyyyMM"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");

                //기준년월T
                oForm.DataSources.UserDataSources.Add("StdYMT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("StdYMT").Specific.DataBind.SetBound(true, "", "StdYMT");
                oForm.DataSources.UserDataSources.Item("StdYMT").Value = DateTime.Now.ToString("yyyyMM"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");
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
        /// 필수 입력사항 check
        /// </summary>
        /// <returns></returns>
        private bool DataValideCheck()
        {
            bool returnValue = false;
            string errCode = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("StdYMF").Specific.VALUE)) //시작월
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("StdYMT").Specific.VALUE)) //종료월
                {
                    errCode = "2";
                    throw new Exception();
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("시작월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("StdYMF").Click(BoCellClickType.ct_Regular);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("종료월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("StdYMT").Click(BoCellClickType.ct_Regular);
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
        /// 재무제표(대차대조표(F01)) 정보 전송
        /// </summary>
        /// <param name="pForm_ID"></param>
        /// <returns></returns>
        private bool DataTransmission(string pForm_ID)
        {
            bool returnValue = false;
            string errCode = string.Empty;
            short loopCount = 0;
            string E_MESSAGE = string.Empty;

            string Query01 = string.Empty;
            string StdYMF = string.Empty;
            string StdYMT = string.Empty;
            string StdDtF = string.Empty;
            string StdDtT = string.Empty;
            string StdDtS = string.Empty; //조회종료년월의 첫일자 저장
            string StdDtN = string.Empty; //조회종료년월의 다음달 저장

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            StdYMF = oForm.Items.Item("StdYMF").Specific.Value.ToString().Trim();
            StdYMT = oForm.Items.Item("StdYMT").Specific.Value.ToString().Trim();

            string Client = null; //클라이언트(운영용:210, 테스트용:710)
            string ServerIP = null; //서버IP(운영용:192.1.11.3, 테스트용:192.1.11.7)

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            RfcDestination rfcDest = null;
            RfcRepository rfcRep = null;
            
            try
            {
                //Real
                Client = "210";
                ServerIP = "192.1.11.3";

                //Test
                //Client = "810"
                //ServerIP = "192.1.11.7"

                //Plant = "9300"

                //0. 연결
                if (dataHelpClass.SAPConnection(ref rfcDest, ref rfcRep, "PSC", ServerIP, Client, "ifuser", "pdauser") == false)
                {
                    errCode = "1";
                    throw new Exception();
                }

                IRfcFunction oFunction = null;
                oFunction = rfcRep.CreateFunction("ZFI_STATEMENT_PSH");
                oFunction.SetValue("I_FORM_ID", pForm_ID); //서식코드(F01:대차대조표)
                oFunction.SetValue("I_SPMONF", StdYMF); //조회시작년월
                oFunction.SetValue("I_SPMONT", StdYMF); //조회종료년월


                //SAPFunctionsOCX.SAPFunctions oSapConnection = new SAPFunctionsOCX.SAPFunctions();
                ////R3연결 객체
                ////oSapConnection = CreateObject("SAP.Functions")
                //oSapConnection.Connection.User = "ifuser";
                //oSapConnection.Connection.Password = "pdauser";
                //oSapConnection.Connection.Client = Client;
                //oSapConnection.Connection.ApplicationServer = ServerIP;
                //oSapConnection.Connection.Language = "KO";
                //oSapConnection.Connection.SystemNumber = "00";

                //if (!oSapConnection.Connection.Logon(0, true))
                //{
                //    ErrNum = 1;
                //    goto DataTransmission_Error;
                //}

                //1. SAP R3 함수 호출(매개변수 전달)
                SAPFunctionsOCX.Function oFunction = new SAPFunctionsOCX.Function();
                //Object 'R3함수용 Function 객체
                oFunction = oSapConnection.Add("ZFI_STATEMENT_PSH");

                oFunction.Exports["I_FORM_ID"] = pForm_ID; //서식코드(F01:대차대조표)
                oFunction.Exports["I_SPMONF"] = StdYMF; //조회시작년월
                oFunction.Exports["I_SPMONT"] = StdYMT; //조회종료년월

                //2. 테이블
                SAPTableFactoryCtrl.Table oTable = null;

                //-------------------------------------------------------------
                // SAP함수 호출한 객체에서 테이블 객체 불러와서 값 할당
                //-------------------------------------------------------------
                oTable = oFunction.Tables("IT_INF");

                //2-1. 조회
                StdDtF = StdYMF + "01";
                StdDtS = Strings.Left(StdYMT, 4) + "-" + Strings.Mid(StdYMT, 5, 2) + "-01"; //조회종료년월의 첫일자
                StdDtN = Convert.ToString(DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Month, 1, Convert.ToDateTime(StdDtS))); //조회종료년월의 다음달
                StdDtT = Strings.Replace(Convert.ToString(DateAndTime.DateSerial(DateAndTime.Year(Convert.ToDateTime(StdDtS)), DateAndTime.Month(Convert.ToDateTime(StdDtN)), 1 - 1)), "-", ""); //조회종료년월의 말일

                if (pForm_ID == "F01") //대차대조표
                {
                    Query01 = "         EXEC PS_CO600_01 '";
                    Query01 = Query01 + StdDtF + "','";
                    Query01 = Query01 + StdDtT + "','";
                    Query01 = Query01 + "T" + "'";
                }
                else if (pForm_ID == "F02") //제조원가명세서
                {
                    Query01 = "         EXEC PS_CO600_02 '";
                    Query01 = Query01 + StdDtF + "','";
                    Query01 = Query01 + StdDtT + "','";
                    Query01 = Query01 + "T" + "'";
                }
                else if (pForm_ID == "F03") //손익계산서
                {
                    Query01 = "         EXEC PS_CO600_04 '";
                    Query01 = Query01 + StdDtF + "','";
                    Query01 = Query01 + StdDtT + "','";
                    Query01 = Query01 + "T" + "'";
                }

                RecordSet01.DoQuery(Query01);

                for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++)
                {
                    oTable.Rows.Add();

                    oTable[loopCount + 1, "NODE_KEY"] = RecordSet01.Fields.Item("AcctCode").Value; //NODE_KEY    '키필드. 계정관리상의 계정코드 (저희쪽 데이터랑 매핑하기 위해 필요합니다)
                    oTable[loopCount + 1, "TITLE1"] = RecordSet01.Fields.Item("Cont1").Value; //TITLE1 '목차제목1
                    oTable[loopCount + 1, "TITLE2"] = RecordSet01.Fields.Item("Cont2").Value; //TITLE2 '목차제목2
                    oTable[loopCount + 1, "AMT_PLANT1"] = RecordSet01.Fields.Item("BPLID_1").Value; //AMT_PLANT1  '창원사업장 금액
                    oTable[loopCount + 1, "AMT_PLANT2"] = RecordSet01.Fields.Item("BPLID_2").Value; //AMT_PLANT2  '안강(사상)사업장 금액
                    oTable[loopCount + 1, "AMT_PLANT3"] = RecordSet01.Fields.Item("BPLID_3").Value; //AMT_PLANT3  '부산사업장 금액
                    oTable[loopCount + 1, "Seq"] = RecordSet01.Fields.Item("Seq").Value; //Seq '목차순서(참고용)

                    RecordSet01.MoveNext();
                }

                //-------------------------------------------------------------
                // 함수를 Call 하면 oTable 객체의 내용을 그대로 SAP R3에 전달.
                // 그러면 SAP R3는 oTable의 각각의 값들에 대해서 처리 결과를 리턴.
                //-------------------------------------------------------------
                if (!oFunction.Call())
                {
                    oSapConnection.Connection.Logoff();
                    ErrNum = 99;
                    goto DataTransmission_Error;
                }
                else
                {
                    E_MESSAGE = Strings.Trim(oFunction.Imports["EV_MSGE"].VALUE); //메시지
                }

                //Call MDC_Com.MDC_GF_Message(E_MESSAGE, "S")

                //oSapConnection.Connection.Logoff(); //R3 연결해제

                //if (ErrNum == 1)
                //{
                //    MDC_Com.MDC_GF_Message(ref "풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.", ref "E");
                //}
                //else if (ErrNum == 99)
                //{
                //    MDC_Com.MDC_GF_Message(ref (oFunction.Exception), ref "E");
                //}
                //else
                //{
                //    MDC_Com.MDC_GF_Message(ref "DataTransmission_Error:" + Err().Number + " - " + Err().Description, ref "E");
                //}

            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

        #region Raise_ItemEvent
        ////****************************************************************************************************************
        ////// ItemEventHander
        ////****************************************************************************************************************
        //		public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;
        //			int ErrNum = 0;
        //			SAPbouiCOM.ProgressBar ProgressBar01 = null;

        //			SAPbouiCOM.ProgressBar ProgBar01 = null;
        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {
        //				switch (pval.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //						////1
        //						if (pval.ItemUID == "1") {
        //							if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //								//                        If DataValideCheck = False Then
        //								//                            BubbleEvent = False
        //								//                            Exit Sub
        //								//                        End If
        //								//                        If MatrixSpaceLineDel = False Then
        //								//                            BubbleEvent = False
        //								//                            Exit Sub
        //								//                        End If
        //							}

        //						//전송 클릭시
        //						} else if (pval.ItemUID == "Btn01") {
        //							if (DataValideCheck() == false) {
        //								BubbleEvent = false;
        //								return;
        //							} else {

        //								ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("전송 중...", 100, false);

        //								//대차대조표
        //								if (DataTransmission("F01") == false) {
        //									ErrNum = 1;
        //									goto Raise_ItemEvent_Error;
        //								}

        //								//제조원가명세서
        //								if (DataTransmission("F02") == false) {
        //									ErrNum = 2;
        //									goto Raise_ItemEvent_Error;
        //								}

        //								//손익계산서
        //								if (DataTransmission("F03") == false) {
        //									ErrNum = 3;
        //									goto Raise_ItemEvent_Error;
        //								}

        //								ProgBar01.Value = 100;
        //								ProgBar01.Stop();
        //								//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //								ProgBar01 = null;

        //								MDC_Com.MDC_GF_Message(ref "전송완료.", ref "W");

        //							}
        //						}
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //						////2
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //						////5
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_CLICK:
        //						////6
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //						////7
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //						////8
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //						////10
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //						////11
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //						////18
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //						////19
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //						////20
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //						////27
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //						////3
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //						////4
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //						////17
        //						break;
        //				}

        //				//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.EventType) {
        //					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //						////1
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //						////2
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //						////5
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_CLICK:
        //						////6
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //						////7
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //						////8
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //						////10
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //						////11
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //						////18
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //						////19
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //						////20
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //						////27
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //						////3
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //						////4
        //						break;
        //					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //						////17
        //						SubMain.RemoveForms(oFormUniqueID01);
        //						//UPGRADE_NOTE: oForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //						oForm01 = null;
        //						break;
        //					//                Set oMat01 = Nothing
        //				}
        //			}
        //			return;
        //			Raise_ItemEvent_Error:

        //			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			ProgressBar01 = null;
        //			if (ErrNum == 101) {
        //				ErrNum = 0;
        //				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //				BubbleEvent = false;
        //			} else if (ErrNum == 1) {
        //				MDC_Com.MDC_GF_Message(ref "대차대조표 전송중 오류발생.", ref "E");
        //			} else if (ErrNum == 2) {
        //				MDC_Com.MDC_GF_Message(ref "제조원가명세서 전송중 오류발생.", ref "E");
        //			} else if (ErrNum == 3) {
        //				MDC_Com.MDC_GF_Message(ref "손익계산서 전송중 오류발생.", ref "E");
        //			} else {
        //				MDC_Com.MDC_GF_Message(ref "Raise_ItemEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //			}
        //		}
        #endregion

        #region Raise_MenuEvent
        //		public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			int i = 0;

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
        //						break;
        //					case "1281":
        //						//찾기
        //						break;
        //					case "1282":
        //						//추가
        //						break;
        //					case "1285":
        //						//복원
        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //				}

        //				//-----------------------------------------------------------------------------------------------------------
        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {
        //				switch (pval.MenuUID) {
        //					case "1284":
        //						//취소
        //						break;
        //					case "1286":
        //						//닫기
        //						break;
        //					case "1285":
        //						//복원
        //						break;
        //					case "1293":
        //						//행삭제
        //						break;
        //					case "1281":
        //						//찾기
        //						break;
        //					case "1282":
        //						//추가
        //						break;
        //					case "1288":
        //					case "1289":
        //					case "1290":
        //					case "1291":
        //						//레코드이동버튼
        //						break;
        //				}
        //			}
        //			return;
        //			Raise_MenuEvent_Error:
        //			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //			MDC_Com.MDC_GF_Message(ref "Raise_MenuEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //		}
        #endregion

        #region Raise_RightClickEvent
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if ((eventInfo.BeforeAction == true)) {

        //			} else if ((eventInfo.BeforeAction == false)) {
        //				////작업
        //			}
        //			return;
        //			Raise_RightClickEvent_Error:

        //			SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

        //			MDC_Com.MDC_GF_Message(ref "Raise_FormDataEvent_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //		}
        #endregion





    }
}
