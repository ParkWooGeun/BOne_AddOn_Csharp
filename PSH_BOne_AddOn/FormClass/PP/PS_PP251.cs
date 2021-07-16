using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 월별 가동률 관리(기계)
	/// </summary>
	internal class PS_PP251 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;

		private SAPbouiCOM.DBDataSource oDS_PS_PP251L;
			
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP251.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}
				oFormUniqueID = "PS_PP251_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP251");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP251_CreateItems();
				PS_PP251_ComboBox_Setting();
				PS_PP251_Initial_Setting();
				PS_PP251_FormResize();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Items.Item("Folder01").Specific.Select(); //폼이 로드 될 때 Folder01이 선택됨
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP251_CreateItems
		/// </summary>
		private void PS_PP251_CreateItems()
		{
			try
			{
				oDS_PS_PP251L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM01").Specific.DataBind.SetBound(true, "", "StdYM01");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType01").Specific.DataBind.SetBound(true, "", "CardType01");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType01").Specific.DataBind.SetBound(true, "", "ItemType01");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 11);
				oForm.Items.Item("OrdNum01").Specific.DataBind.SetBound(true, "", "OrdNum01");

				//품명
				oForm.DataSources.UserDataSources.Add("OrdName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("OrdName01").Specific.DataBind.SetBound(true, "", "OrdName01");

				//규격
				oForm.DataSources.UserDataSources.Add("OrdSpec01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("OrdSpec01").Specific.DataBind.SetBound(true, "", "OrdSpec01");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP251_ComboBox_Setting
		/// </summary>
		private void PS_PP251_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//거래처구분
				sQry = " SELECT      U_Minor,";
				sQry += "             U_CdName";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'C100'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Seq";
				oForm.Items.Item("CardType01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType01").Specific, sQry, "%", false, false);

				//품목구분
				sQry = " SELECT      U_Minor,";
				sQry += "             U_CdName";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'S002'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Seq";
				oForm.Items.Item("ItemType01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType01").Specific, sQry, "%", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP251_Initial_Setting
		/// </summary>
		private void PS_PP251_Initial_Setting()
		{
			try
			{
				//기준일자
				oForm.Items.Item("StdYM01").Specific.Value = DateTime.Now.ToString("yyyyMM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP251_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="prmMat"></param>
		/// <param name="prmDataSource"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP251_AddMatrixRow(int oRow, SAPbouiCOM.Matrix prmMat, SAPbouiCOM.DBDataSource prmDataSource, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);

				if (RowIserted == false) //행추가여부
				{
					prmDataSource.InsertRecord(oRow);
				}
				prmMat.AddRow();
				prmDataSource.Offset = oRow;
				prmMat.LoadFromDataSource();
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
		/// PS_PP251_FormResize
		/// </summary>
		private void PS_PP251_FormResize()
		{
			try
			{
				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Mat01").Height + 120;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Mat01").Width + 20;
				oMat01.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP251_MTX01 가동률 관리 조회
		/// </summary>
		private void PS_PP251_MTX01()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;

			string StdYM;	 //기준년월
			string CardType; //거래처구분
			string ItemType; //품목구분
			string OrdNum;   //작번

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				StdYM    = oForm.Items.Item("StdYM01").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType01").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType01").Specific.Value.ToString().Trim();
				OrdNum   = oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = " EXEC PS_PP251_01 '";
				sQry += StdYM + "','";
				sQry += CardType + "','";
				sQry += ItemType + "','";
				sQry += OrdNum + "'";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat01.Clear();
					errMessage = "조회 결과가 없습니다. 확인하세요.:";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_PP251L.InsertRecord(loopCount);
					}
					oDS_PS_PP251L.Offset = loopCount;

					oDS_PS_PP251L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));					          //라인번호
					oDS_PS_PP251L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("Select").Value.ToString().Trim());	  //선택
					oDS_PS_PP251L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());  //품목코드
					oDS_PS_PP251L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());  //거래처명
					oDS_PS_PP251L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("InOut").Value.ToString().Trim());	  //자체 / 외주
					oDS_PS_PP251L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("GrpName").Value.ToString().Trim());	  //계열사명
					oDS_PS_PP251L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());  //품명
					oDS_PS_PP251L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim());  //규격
					oDS_PS_PP251L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("ItemUnit").Value.ToString().Trim());  //단위
					oDS_PS_PP251L.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("WOQty").Value.ToString().Trim());	  //지시수량
					oDS_PS_PP251L.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("WorkQty").Value.ToString().Trim());	  //생산수량
					oDS_PS_PP251L.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("OrdQty").Value.ToString().Trim());	  //수주수량
					oDS_PS_PP251L.SetValue("U_ColSum04", loopCount, oRecordSet.Fields.Item("OrdAmt").Value.ToString().Trim());	  //수주금액
					oDS_PS_PP251L.SetValue("U_ColQty01", loopCount, oRecordSet.Fields.Item("StdTime").Value.ToString().Trim());	  //표준공수
					oDS_PS_PP251L.SetValue("U_ColQty02", loopCount, oRecordSet.Fields.Item("WorkTime").Value.ToString().Trim());  //실동공수(전월까지)
					oDS_PS_PP251L.SetValue("U_ColQty03", loopCount, oRecordSet.Fields.Item("AddTime").Value.ToString().Trim());	  //추가예상공수
					oDS_PS_PP251L.SetValue("U_ColQty04", loopCount, oRecordSet.Fields.Item("TotTime").Value.ToString().Trim());	  //공수합계
					oDS_PS_PP251L.SetValue("U_ColQty05", loopCount, oRecordSet.Fields.Item("Plan1st").Value.ToString().Trim());	  //1주계획
					oDS_PS_PP251L.SetValue("U_ColQty06", loopCount, oRecordSet.Fields.Item("Rslt1st").Value.ToString().Trim());	  //1주실적
					oDS_PS_PP251L.SetValue("U_ColQty07", loopCount, oRecordSet.Fields.Item("Plan2nd").Value.ToString().Trim());	  //2주계획
					oDS_PS_PP251L.SetValue("U_ColQty08", loopCount, oRecordSet.Fields.Item("Rslt2nd").Value.ToString().Trim());	  //2주실적
					oDS_PS_PP251L.SetValue("U_ColQty09", loopCount, oRecordSet.Fields.Item("Plan3rd").Value.ToString().Trim());	  //3주계획
					oDS_PS_PP251L.SetValue("U_ColQty10", loopCount, oRecordSet.Fields.Item("Rslt3rd").Value.ToString().Trim());	  //3주실적
					oDS_PS_PP251L.SetValue("U_ColQty11", loopCount, oRecordSet.Fields.Item("Plan4th").Value.ToString().Trim());	  //4주계획
					oDS_PS_PP251L.SetValue("U_ColQty12", loopCount, oRecordSet.Fields.Item("Rslt4th").Value.ToString().Trim());	  //4주실적
					oDS_PS_PP251L.SetValue("U_ColQty13", loopCount, oRecordSet.Fields.Item("Plan5th").Value.ToString().Trim());	  //5주계획
					oDS_PS_PP251L.SetValue("U_ColQty14", loopCount, oRecordSet.Fields.Item("Rslt5th").Value.ToString().Trim());	  //5주실적
					oDS_PS_PP251L.SetValue("U_ColQty15", loopCount, oRecordSet.Fields.Item("PlanNMth").Value.ToString().Trim());  //차월계획
					oDS_PS_PP251L.SetValue("U_ColQty16", loopCount, oRecordSet.Fields.Item("PlanNNMth").Value.ToString().Trim()); //차차월계획
					oDS_PS_PP251L.SetValue("U_ColQty17", loopCount, oRecordSet.Fields.Item("PlanLMth").Value.ToString().Trim());  //이후계획
					oDS_PS_PP251L.SetValue("U_ColDt01", loopCount, Convert.ToDateTime(oRecordSet.Fields.Item("RegDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //최종등록일

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
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
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP251_SaveData01
		/// </summary>
		private void PS_PP251_SaveData01()
		{
			int loopCount;
			string sQry;

			string StdYM;	  //기준년월
			string OrdNum;	  //품목코드
			double AddTime;	  //추가예상공수
			double Plan1st;	  //1주차계획
			double Plan2nd;	  //2주차계획
			double Plan3rd;	  //3주차계획
			double Plan4th;	  //4주차계획
			double Plan5th;	  //5주차계획
			double PlanNMth;  //차월계획
			double PlanNNMth; //차차월계획
			double PlanLMth;  //이후계획
			string RegDate;   //최종등록일

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			try
			{
				StdYM = oForm.DataSources.UserDataSources.Item("StdYM01").Value.ToString().Trim();

				oMat01.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP251L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						OrdNum = oDS_PS_PP251L.GetValue("U_ColReg02", loopCount).ToString().Trim();	                     //품목코드
						AddTime = Convert.ToDouble(oDS_PS_PP251L.GetValue("U_ColQty03", loopCount).ToString().Trim());   //추가예상공수
						Plan1st = Convert.ToDouble(oDS_PS_PP251L.GetValue("U_ColQty05", loopCount).ToString().Trim());   //1주차계획
						Plan2nd = Convert.ToDouble(oDS_PS_PP251L.GetValue("U_ColQty07", loopCount).ToString().Trim());   //2주차계획
						Plan3rd = Convert.ToDouble(oDS_PS_PP251L.GetValue("U_ColQty09", loopCount).ToString().Trim());   //3주차계획
						Plan4th = Convert.ToDouble(oDS_PS_PP251L.GetValue("U_ColQty11", loopCount).ToString().Trim());   //4주차계획
						Plan5th = Convert.ToDouble(oDS_PS_PP251L.GetValue("U_ColQty13", loopCount).ToString().Trim());   //5주차계획
						PlanNMth = Convert.ToDouble(oDS_PS_PP251L.GetValue("U_ColQty15", loopCount).ToString().Trim());	 //차월계획
						PlanNNMth = Convert.ToDouble(oDS_PS_PP251L.GetValue("U_ColQty16", loopCount).ToString().Trim()); //차차월계획
						PlanLMth = Convert.ToDouble(oDS_PS_PP251L.GetValue("U_ColQty17", loopCount).ToString().Trim());  //이후계획
  					    RegDate = oDS_PS_PP251L.GetValue("U_ColDt01", loopCount).ToString().Trim();                      //최종등록일

						sQry = " EXEC [PS_PP251_02] '";
						sQry += StdYM + "','";
						sQry += OrdNum + "','";
						sQry += AddTime + "','";
						sQry += Plan1st + "','";
						sQry += Plan2nd + "','";
						sQry += Plan3rd + "','";
						sQry += Plan4th + "','";
						sQry += Plan5th + "','";
						sQry += PlanNMth + "','";
						sQry += PlanNNMth + "','";
						sQry += PlanLMth + "','";
						sQry += RegDate + "'";
						oRecordSet.DoQuery(sQry);

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToString(oMat01.VisualRowCount - 1) + "건 저장중...";
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP251_CheckAll01 체크박스 전체선택, 해제
		/// </summary>
		private void PS_PP251_CheckAll01()
		{
			string CheckType;
			int loopCount;

			try
			{
				oForm.Freeze(true);

				CheckType = "Y";
				oMat01.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP251L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break; // TODO: might not be correct. Was : Exit For
					}
				}

				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
				{
					oDS_PS_PP251L.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_PP251L.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_PP251L.SetValue("U_ColReg01", loopCount, "N");
					}
				}

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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "BtnSrch01")
					{
						PS_PP251_MTX01(); //매트릭스에 데이터 로드
					}
					else if (pVal.ItemUID == "BtnSave01")
					{
						PS_PP251_SaveData01();
					}
					else if (pVal.ItemUID == "BtnCheck01")
					{
						PS_PP251_CheckAll01();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스
					if (pVal.ItemUID == "Folder01")
					{
						oForm.Freeze(true);
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
						oForm.Settings.MatrixUID = "Mat01";
						oMat01.AutoResizeColumns();
						oForm.Freeze(false);
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum01", "");
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat01.SelectRow(pVal.Row, true, false);
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
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
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
						if (pVal.ItemUID == "OrdNum01")
						{
							oForm.Items.Item("OrdName01").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "OITM", "'" + oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim() + "'", "");
							oForm.Items.Item("OrdSpec01").Specific.Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "OITM", "'" + oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim() + "'", "");
						}
						oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
					PS_PP251_FormResize();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_GOT_FOCUS
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
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
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP251L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_RightClickEvent
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
									 //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							PS_PP251_AddMatrixRow(oMat01.VisualRowCount, oMat01, oDS_PS_PP251L, false);
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
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "1287": //복제
							break;
						case "7169": //엑셀 내보내기
									 //엑셀 내보내기 이후 처리
							oForm.Freeze(true);
							PS_PP251_AddMatrixRow(oMat01.VisualRowCount, oMat01, oDS_PS_PP251L, false);
							oMat01.LoadFromDataSource();
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
		}
	}
}
