using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 결과 승인 및 전송
	/// </summary>
	internal class PS_QM702 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.Grid oGrid1;
		private SAPbouiCOM.DBDataSource oDS_PS_QM702A;
		private SAPbouiCOM.DBDataSource oDS_PS_QM702B;
		private string oLastItemUID;
		private string oLastColUID;
		private int oLastColRow;



		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM702.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM702_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM702");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM702_CreateItems();
				PS_QM702_FormReset();
				PS_QM702_ComboBox_Setting();
				PS_QM702_EnableMenus();
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
		/// PS_QM702_CreateItems
		/// </summary>
		private void PS_QM702_CreateItems()
		{
			try
			{
				oDS_PS_QM702A = oForm.DataSources.DBDataSources.Item(" ");

				//oGrid1 = oForm.Items.Item("Grid1").Specific;

				oMat01 = oForm.Items.Item("oMat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				//oMat02 = oForm.Items.Item("oMat02").Specific;
				//oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				//oMat02.AutoResizeColumns();


				////관리번호
				//oForm.DataSources.UserDataSources.Add("SDocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				//oForm.Items.Item("SDocEntry").Specific.DataBind.SetBound(true, "", "SDocEntry");

				////사업장
				//oForm.DataSources.UserDataSources.Add("SCLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				//oForm.Items.Item("SCLTCOD").Specific.DataBind.SetBound(true, "", "SCLTCOD");

				////작번
				//oForm.DataSources.UserDataSources.Add("SItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				//oForm.Items.Item("SItemCode").Specific.DataBind.SetBound(true, "", "SItemCode");

				////기간(FR)
				//oForm.DataSources.UserDataSources.Add("SFrDate", SAPbouiCOM.BoDataType.dt_DATE);
				//oForm.Items.Item("SFrDate").Specific.DataBind.SetBound(true, "", "SFrDate");

				////기간(TO)
				//oForm.DataSources.UserDataSources.Add("SToDate", SAPbouiCOM.BoDataType.dt_DATE);
				//oForm.Items.Item("SToDate").Specific.DataBind.SetBound(true, "", "SToDate");

				//oMat.Columns.Item("Check").Visible = false; //선택 체크박스 Visible = False

				////SET
				//oForm.Items.Item("SFrDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
				//oForm.Items.Item("SToDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 콤보박스 set
		/// </summary>
		private void PS_QM702_ComboBox_Setting()
		{
			string User_BPLId;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				User_BPLId = dataHelpClass.User_BPLID();
				//조회정보 사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("SCLTCOD").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", User_BPLId, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM702_EnableMenus
		/// </summary>
		private void PS_QM702_EnableMenus()
		{
			try
			{
				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1285", false); //복원
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1293", false); //행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
				dataHelpClass.SetEnableMenus(oForm, false, false, false, false, false, true, true, true, true, true, false, false, false, false, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 메트릭스 Row추가
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM702_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_QM702L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM702L.Offset = oRow;
				oDS_PS_QM702L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}
		/// <summary>
		/// 데이터 조회
		/// </summary>
		private void PS_QM702_FormReset()
		{
			string sQry;
			string errMessange = string.Empty;
			SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			try
			{
				sQry = "EXEC [PS_QM702_01] '" + dataHelpClass.User_MSTCOD() "'";
				oRecordSet01.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_QM702A.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet01.RecordCount == 0)
				{
					errMessage = "예약내역이 존재하지 않습니다. 등록을 진행하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_QM702A.Size)
					{
						oDS_PS_QM702A.InsertRecord(i);
					}

					oMat01.AddRow();
					oDS_PS_QM702A.Offset = i;

					oDS_PS_QM702A.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM702A.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("U_InOut").Value.ToString().Trim()); //구분
					oDS_PS_QM702A.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim()); //문서번호
					oDS_PS_QM702A.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("U_CLTCOD").Value.ToString().Trim()); //사업장
					oDS_PS_QM702A.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("U_WorkNum").Value.ToString().Trim()); //작업번호
					oDS_PS_QM702A.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim()); //품목명
					oDS_PS_QM702A.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("U_WorkDate").Value.ToString().Trim()); //검사일자
					oDS_PS_QM702A.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("U_WorkCode").Value.ToString().Trim()); //검사자
					oDS_PS_QM702A.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("U_WorkName").Value.ToString().Trim()); //검사자명
					oDS_PS_QM702A.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim()); //거래처코드
					oDS_PS_QM702A.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("U_CardName").Value.ToString().Trim()); //거래처명
					oDS_PS_QM702A.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("U_BZZadQty").Value.ToString().Trim()); //불량수량
					oDS_PS_QM702A.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("U_BadCode").Value.ToString().Trim()); //불량분류
					oDS_PS_QM702A.SetValue("U_ColReg13", i, oRecordSet01.Fields.Item("U_BadNote").Value.ToString().Trim()); //불량원인
					oDS_PS_QM702A.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("U_verdict").Value.ToString().Trim()); //판정의견
					oDS_PS_QM702A.SetValue("U_ColReg14", i, oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim()); //근거
					oRecordSet01.MoveNext();
					ProgBar01.Value += 1;
					ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
				}
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				ProgBar01.Stop();
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
			}
        }

		///// <summary>
		///// 데이터 조회
		///// </summary>
		////private void PS_QM702_MTX01()
  //      {
  //          string sQry;
  //          string errMessange = string.Empty;
  //          SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
  //          SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

  //          try
  //          {
  //              sQry = "EXEC [PH_PY011_01] '" + dataHelpClass.User_MSTCOD() "'";
  //              oRecordSet01.DoQuery(sQry);

  //              oMat01.Clear();
  //              oDS_PS_QM702A.Clear();
  //              oMat01.FlushToDataSource();
  //              oMat01.LoadFromDataSource();

  //              if (oRecordSet01.RecordCount == 0)
  //              {
  //                  errMessage = "예약내역이 존재하지 않습니다. 등록을 진행하세요.";
  //                  throw new Exception();
  //              }

  //              for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
  //              {
  //                  if (i + 1 > oDS_PS_QM702A.Size)
  //                  {
  //                      oDS_PS_QM702A.InsertRecord(i);
  //                  }

  //                  oMat01.AddRow();
  //                  oDS_PS_QM702A.Offset = i;

  //                  oDS_PS_QM702A.SetValue("U_LineNum", i, Convert.ToString(i + 1));
  //                  oDS_PS_QM702A.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim()); //사번
  //                  oDS_PS_QM702A.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("FullName").Value.ToString().Trim()); //성명
  //                  oDS_PS_QM702A.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("GrpDat").Value.ToString().Trim()); //입사일자
  //                  oDS_PS_QM702A.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("gunsok").Value.ToString().Trim()); //근속년수
  //                  oDS_PS_QM702A.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("Callname").Value.ToString().Trim()); //현재코드
  //                  oDS_PS_QM702A.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("Cname").Value.ToString().Trim()); //현재호칭
  //                  oDS_PS_QM702A.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("ChCallName").Value.ToString().Trim()); //변경후코드
  //                  oDS_PS_QM702A.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("ChCName").Value.ToString().Trim()); //변경후호칭
  //                  oDS_PS_QM702A.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("ChYN").Value.ToString().Trim()); //변경여부

  //                  if (oRecordSet01.Fields.Item("ChYN").Value.Trim() == "Y")
  //                  {
  //                      ChCnt = (short)(ChCnt + 1);
  //                  }

  //                  oRecordSet01.MoveNext();
  //                  ProgBar01.Value += 1;
  //                  ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

  //              }

  //              oMat01.LoadFromDataSource();
  //              oMat01.AutoResizeColumns();
  //              ProgBar01.Stop();

  //              oForm.Items.Item("ChCnt").Specific.Value = ChCnt;
  //          }
  //          catch (Exception ex)
  //          {
  //              ProgBar01.Stop();
  //          }
  //          finally
  //          {
  //              oForm.Freeze(false);
  //              if (ProgBar01 != null)
  //              {
  //                  ProgBar01.Stop();
  //                  System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
  //              }
  //              System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
  //          }
  //      }

        /// <summary>
        /// 데이터 UPDATE
        /// </summary>
        /// <returns></returns>
        private bool PS_QM702_UpdateData()
		{
			bool ReturnValue = false;
			string DocEntry;    //관리번호
			string CLTCOD;   //사업장
			string ItemCode; //작번
			string ItemName; //품명
			string ItemSpec; //규격
			string DocDate;  //날짜
			decimal TotalQty;    //전체수량
			decimal BadQty;      //불량수량
			string BadNote;  //불량내용
			string CpCode;   //원인공정
			string CpName;   //원인공정명
			string WorkCode; //작업자
			string WorkName; //작업자명
			string LastNote; //최종판정
			decimal Cost01;  //자재
			decimal Cost02;  //가공
			decimal Cost03;  //설계
			decimal Cost04;  //외주
			decimal Cost05;  //분해조립
			decimal Cost06;  //A/S출장
			decimal Cost07;  //운송
			decimal Cost08;  //지체상금
			decimal CostTot; //계
			string Comments; //비고
			string UserSign; //UserSign
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				ItemSpec = oForm.Items.Item("ItemSpec").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				TotalQty = Convert.ToDecimal(oForm.Items.Item("TotalQty").Specific.Value.ToString().Trim());
				BadQty = Convert.ToDecimal(oForm.Items.Item("BadQty").Specific.Value.ToString().Trim());
				BadNote = oForm.Items.Item("BadNote").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				CpName = oForm.Items.Item("CpName").Specific.Value.ToString().Trim();
				WorkCode = oForm.Items.Item("WorkCode").Specific.Value.ToString().Trim();
				WorkName = oForm.Items.Item("WorkName").Specific.Value.ToString().Trim();
				LastNote = oForm.Items.Item("LastNote").Specific.Value.ToString().Trim();
				UserSign = Convert.ToString(PSH_Globals.oCompany.UserSignature);

				sQry = " EXEC [PS_QM702_03] ";
				sQry += "'" + DocEntry + "',";
				sQry += "'" + CLTCOD + "',";
				sQry += "'" + ItemCode + "',";
				sQry += "'" + ItemName + "',";
				sQry += "'" + ItemSpec + "',";
				sQry += "'" + DocDate + "',";
				sQry += "'" + TotalQty + "',";
				sQry += "'" + BadQty + "',";
				sQry += "'" + BadNote + "',";
				sQry += "'" + CpCode + "',";
				sQry += "'" + CpName + "',";
				sQry += "'" + WorkCode + "',";
				sQry += "'" + WorkName + "',";
				sQry += "'" + LastNote + "',";
		        sQry += "'" + UserSign + "'";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.StatusBar.SetText("수정 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				ReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
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
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
				//    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
				//    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
				//    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
				//    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
				//    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
					Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
					break;
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
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnAdd")
					{
					}
					else if (pVal.ItemUID == "BtnSearch")
					{
						PS_QM702_FormReset();
						PS_QM702_MTX01();
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
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						//if (pVal.Row > 0 && !string.IsNullOrEmpty(oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
						//{
						//	oMat.SelectRow(pVal.Row, true, false);
						//	//DataSource를 이용하여 각 컨트롤에 값을 출력
						//	oDS_PS_QM702H.SetValue("DocEntry", 0, oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//관리번호
						//	oDS_PS_QM702H.SetValue("U_CLTCOD", 0, oMat.Columns.Item("CLTCOD").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//사업장
						//	oDS_PS_QM702H.SetValue("U_ItemCode", 0, oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//작번
						//	oDS_PS_QM702H.SetValue("U_ItemName", 0, oMat.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//품명
						//	oDS_PS_QM702H.SetValue("U_ItemSpec", 0, oMat.Columns.Item("ItemSpec").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//규격
						//	oDS_PS_QM702H.SetValue("U_DocDate", 0, oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//날짜
						//	oDS_PS_QM702H.SetValue("U_TotalQty", 0, oMat.Columns.Item("TotalQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//전체수량
						//	oDS_PS_QM702H.SetValue("U_BadQty", 0, oMat.Columns.Item("BadQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//불량수량
						//	oDS_PS_QM702H.SetValue("U_BadNote", 0, oMat.Columns.Item("BadNote").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//불량내용
						//	oDS_PS_QM702H.SetValue("U_CpCode", 0, oMat.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());		//원인공정
						//	oDS_PS_QM702H.SetValue("U_CpName", 0, oMat.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());     //원인공정명
						//	oDS_PS_QM702H.SetValue("U_WorkCode", 0, oMat.Columns.Item("WorkCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//작업자
						//	oDS_PS_QM702H.SetValue("U_WorkName", 0, oMat.Columns.Item("WorkName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());	//작업자명
						//	oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						//	PS_QM702_LoadCaption();
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_QM702_FlushToItemValue(pVal.ItemUID);
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
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_QM702_FormResize();
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
						}
						else
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM702H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM702L);
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
						case "7169": //엑셀 내보내기
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
							break;
						case "1282": //추가
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
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
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}
	}
}
