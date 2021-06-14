using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 기계공구,몰드 품목코드 등록
	/// </summary>
	internal class PS_SD001 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private string oCardCode;
		private string oCpNaming;
		private string oWhsCode;
		private string oItmBsort;
		private string oItmMsort;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD001.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD001_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD001");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				CreateItems();
				ComboBox_Setting();
												   
				oForm.EnableMenu("1281", false); // 찾기
				oForm.EnableMenu("1282", true);  // 추가
				oForm.EnableMenu("1283", true);  // 제거
				oForm.EnableMenu("1287", true);  // 복제
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1288", false);
				oForm.EnableMenu("1289", false);
				oForm.EnableMenu("1290", false);
				oForm.EnableMenu("1291", false);
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
		/// CreateItems
		/// </summary>
		private void CreateItems()
		{
			try
			{
				oForm.DataSources.UserDataSources.Add("MatDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("MatDate").Specific.DataBind.SetBound(true, "", "MatDate");
				oForm.DataSources.UserDataSources.Item("MatDate").Value = DateTime.Now.ToString("yyyyMMdd");

				// Radio Button 처리
				oForm.DataSources.UserDataSources.Add("RadioBtn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

				// 영업
				oForm.Items.Item("Ra_Sale").Specific.ValOn = "A";
				oForm.Items.Item("Ra_Sale").Specific.ValOff = "0";
				oForm.Items.Item("Ra_Sale").Specific.DataBind.SetBound(true, "", "RadioBtn");
				oForm.Items.Item("Ra_Sale").Specific.Selected = true;

				// 견본
				oForm.Items.Item("Ra_Samp").Specific.ValOn = "B";
				oForm.Items.Item("Ra_Samp").Specific.ValOff = "0";
				oForm.Items.Item("Ra_Samp").Specific.DataBind.SetBound(true, "", "RadioBtn");
				oForm.Items.Item("Ra_Samp").Specific.GroupWith("Ra_Sale");

				// AS
				oForm.Items.Item("Ra_AS").Specific.ValOn = "C";
				oForm.Items.Item("Ra_AS").Specific.ValOff = "0";
				oForm.Items.Item("Ra_AS").Specific.DataBind.SetBound(true, "", "RadioBtn");
				oForm.Items.Item("Ra_AS").Specific.GroupWith("Ra_Sale");

				// 멀티
				oForm.Items.Item("Ra_Multi").Specific.ValOn = "D";
				oForm.Items.Item("Ra_Multi").Specific.ValOff = "0";
				oForm.Items.Item("Ra_Multi").Specific.DataBind.SetBound(true, "", "RadioBtn");
				oForm.Items.Item("Ra_Multi").Specific.GroupWith("Ra_Sale");

				// 신동
				oForm.Items.Item("Ra_Sin").Specific.ValOn = "E";
				oForm.Items.Item("Ra_Sin").Specific.ValOff = "0";
				oForm.Items.Item("Ra_Sin").Specific.DataBind.SetBound(true, "", "RadioBtn");
				oForm.Items.Item("Ra_Sin").Specific.GroupWith("Ra_Sale");

				//R&D(설계)
				oForm.Items.Item("Ra_RND").Specific.ValOn = "R";
				oForm.Items.Item("Ra_RND").Specific.ValOff = "0";
				oForm.Items.Item("Ra_RND").Specific.DataBind.SetBound(true, "", "RadioBtn");
				oForm.Items.Item("Ra_RND").Specific.GroupWith("Ra_Sale");

				// 품목구분
				oForm.Items.Item("CpNaming").Specific.ValidValues.Add("D", "임가공");
				oForm.Items.Item("CpNaming").Specific.ValidValues.Add("G", "게이지");
				oForm.Items.Item("CpNaming").Specific.ValidValues.Add("J", "몰드");
				oForm.Items.Item("CpNaming").Specific.ValidValues.Add("M", "장비");
				oForm.Items.Item("CpNaming").Specific.ValidValues.Add("P", "부품");
				oForm.Items.Item("CpNaming").Specific.ValidValues.Add("T", "공구");

				oForm.Items.Item("WhsCode").Specific.Value = "102";
				oForm.Items.Item("WhsName").Specific.Value = "동래";

				//기준작번
				oForm.DataSources.UserDataSources.Add("BaseCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("BaseCode").Specific.DataBind.SetBound(true, "", "BaseCode");

				//기준작명
				oForm.DataSources.UserDataSources.Add("BaseName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("BaseName").Specific.DataBind.SetBound(true, "", "BaseName");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		public void ComboBox_Setting()
		{
			int loopCount;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
				{
					for (loopCount = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
					{
						oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
					}
				}

				oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "선택");
				sQry = " SELECT      U_Minor AS [Code],";
				sQry += "             U_CdName As [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'I001'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Minor";
				dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
				oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//연간품여부
				oForm.Items.Item("YearPdYN").Specific.ValidValues.Add("%", "선택");
				oForm.Items.Item("YearPdYN").Specific.ValidValues.Add("Y", "Y");
				oForm.Items.Item("YearPdYN").Specific.ValidValues.Add("N", "N");
				oForm.Items.Item("YearPdYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//세분화항목
				oForm.Items.Item("Detail").Specific.ValidValues.Add("%", "선택");
				sQry = " SELECT      U_Minor AS [Code],";
				sQry += "             U_CdName As [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'P213'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Minor";
				dataHelpClass.Set_ComboList(oForm.Items.Item("Detail").Specific, sQry, "", false, false);
				oForm.Items.Item("Detail").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// FormItem_Clear
		/// </summary>
		private void FormItem_Clear()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("ItemName").Specific.Value = "";
					oForm.Items.Item("Spec1").Specific.Value = "";
					oForm.Items.Item("Unit1").Specific.Value = "";
					oForm.Items.Item("ItemCode").Specific.Value = "";
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
		/// HeaderSpaceLineDel
		/// </summary>
		/// <param name="ItemUID"></param>
		/// <returns></returns>
		private bool HeaderSpaceLineDel(string ItemUID)
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				if (oForm.DataSources.UserDataSources.Item("RadioBtn").Value.ToString().Trim() == "A"
					&& string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "영업일 경우 고객코드는 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("CpNaming").Specific.Value.ToString().Trim()))
				{
					errMessage = "품목구분은 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("WhsCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "기본창고는 필수입력 사항입니다.확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ItemName").Specific.Value.ToString().Trim()))
				{
					errMessage = "품목이름은 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim()))
				{
					errMessage = "품목대분류는 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ItmMsort").Specific.Value.ToString().Trim()))
				{
					errMessage = "품목중분류는 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("Spec1").Specific.Value.ToString().Trim()))
				{
					errMessage = "규격은 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("Unit1").Specific.Value.ToString().Trim()))
				{
					errMessage = "단위는 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("MatDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "작명일자는 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				
				if (ItemUID == "Btn02")
                {
					if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
					{
						errMessage = "품목코드생성 버튼을 먼저 누르세요. 확인하세요.";
						throw new Exception();
					}

					if (string.IsNullOrEmpty(oForm.Items.Item("BaseCode").Specific.Value.ToString().Trim()))
					{
						if (codeHelpClass.Left(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim(), 1) == "Z")
						{
							errMessage = "서비스작번(Z)은 필히 기준작번을 등록하여야 합니다.";
							throw new Exception();
						}
					}
				}
				
				functionReturnValue = true;
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
			return functionReturnValue;
		}

		/// <summary>
		/// Create_Itemcode
		/// </summary>
		/// <returns></returns>
		private bool Create_Itemcode()
		{
			bool functionReturnValue = false;

			//SAPbobsCOM.Items oItem01;
			int ErrCode;
			string ErrMsg;
			string ItemName;
			int RetVal;

			SAPbobsCOM.Items oItem01 = null;
			//SAPbobsCOM.Items oItem01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oItem01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

				ProgressBar01.Text = "저장 중...";

				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}

				PSH_Globals.oCompany.StartTransaction();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim() + " " + oForm.Items.Item("Spec1").Specific.Value.ToString().Trim();

				oItem01.ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim(); //아이템코드
				oItem01.ItemName = ItemName; //아이템이름
				oItem01.ForeignName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				oItem01.DefaultWarehouse = oForm.Items.Item("WhsCode").Specific.Value.ToString().Trim(); //기본창고
				oItem01.ItemsGroupCode = Convert.ToInt32("102"); //품목그룹
				oItem01.UserFields.Fields.Item("U_ItmBsort").Value = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim();	//품목대분류
				oItem01.UserFields.Fields.Item("U_ItmMsort").Value = oForm.Items.Item("ItmMsort").Specific.Value.ToString().Trim();	//품목중분류
				oItem01.UserFields.Fields.Item("U_Spec1").Value = oForm.Items.Item("Spec1").Specific.Value.ToString().Trim();		//규격
				oItem01.UserFields.Fields.Item("U_Size").Value = oForm.Items.Item("Spec1").Specific.Value.ToString().Trim();		//사이즈
				oItem01.PurchaseUnit = oForm.Items.Item("Unit1").Specific.Value.ToString().Trim();				                    //구매처표현단위
				oItem01.SalesUnit = oForm.Items.Item("Unit1").Specific.Value.ToString().Trim();				                        //판매처표현단위
				oItem01.UserFields.Fields.Item("U_ObasUnit").Value = "101";		//매입기준단위
				oItem01.UserFields.Fields.Item("U_SbasUnit").Value = "101";		//판매기준단위

				oItem01.UserFields.Fields.Item("U_TeamCode").Value = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();	//고객소속
				oItem01.UserFields.Fields.Item("U_YearPdYN").Value = oForm.Items.Item("YearPdYN").Specific.Value.ToString().Trim();	//연간품여부
				oItem01.UserFields.Fields.Item("U_Detail").Value = oForm.Items.Item("Detail").Specific.Value.ToString().Trim();		//세분화항목
				oItem01.UserFields.Fields.Item("U_BaseCode").Value = oForm.Items.Item("BaseCode").Specific.Value.ToString().Trim();	//기준작번(서비스작번(Z) 일 경우 필수 입력)

				RetVal = oItem01.Add();

				if (0 != RetVal)
				{
					PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
					throw new Exception();
				}

				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				}
				oCardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				oCpNaming = oForm.Items.Item("CpNaming").Specific.Value.ToString().Trim();
				oWhsCode = oForm.Items.Item("WhsCode").Specific.Value.ToString().Trim();
				oItmBsort = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim();
				oItmMsort = oForm.Items.Item("ItmMsort").Specific.Value.ToString().Trim();

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				ProgressBar01.Stop();
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem01);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// Create_Form_Code
		/// </summary>
		private void Create_Form_Code()
		{
			string errMessage = string.Empty;
			string sQry;
			string sTxt = string.Empty;
			string sCho;
			string sSeq;
			string sDate;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oForm.DataSources.UserDataSources.Item("RadioBtn").Value.ToString().Trim() == "A")
				{
					sQry = "SELECT U_CdNaming FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);

					if (string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
					{
						errMessage = "거래처 마스터 데이터에 생성코드가 누락되었습니다.";
						throw new Exception();
					}
					else
					{
						sTxt = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					}

				}
				else if (oForm.DataSources.UserDataSources.Item("RadioBtn").Value.ToString().Trim() == "B")
				{
					sTxt = "S";
				}
				else if (oForm.DataSources.UserDataSources.Item("RadioBtn").Value.ToString().Trim() == "C")
				{
					sTxt = "Z";
				}
				else if (oForm.DataSources.UserDataSources.Item("RadioBtn").Value.ToString().Trim() == "D")
				{
					sTxt = "F";
				}
				else if (oForm.DataSources.UserDataSources.Item("RadioBtn").Value.ToString().Trim() == "E")
				{
					sTxt = "E";
				}
				else if (oForm.DataSources.UserDataSources.Item("RadioBtn").Value.ToString().Trim() == "R")
				{
					sTxt = "R";
				}

				sDate = oForm.Items.Item("MatDate").Specific.Value.ToString().Trim();
				sDate = sDate.Substring(0, 6);
				sCho = sTxt + oForm.Items.Item("CpNaming").Specific.Value.ToString().Trim();

				sQry = "SELECT MAX(SubString(ItemCode,3,9)) FROM OITM WHERE SubString(ItemCode,1,2) = '" + sCho + "' AND ";
				sQry = sQry + "SubString(ItemCode,3,6) = '" + sDate + "'";
				oRecordSet.DoQuery(sQry);

				if (string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
				{
					oForm.Items.Item("ItemCode").Specific.Value = sTxt + oForm.Items.Item("CpNaming").Specific.Value.ToString().Trim() + sDate + "001";
				}
				else
				{
					sSeq = Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
					oForm.Items.Item("ItemCode").Specific.Value = sTxt + oForm.Items.Item("CpNaming").Specific.Value.ToString().Trim() + sSeq;
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
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
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
					if (pVal.ItemUID == "Btn02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (HeaderSpaceLineDel(pVal.ItemUID) == false)
							{
								BubbleEvent = false;
								return;
							}

							if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
							{
								if (Create_Itemcode() == true)
								{
									PSH_Globals.SBO_Application.StatusBar.SetText("신규 자재 등록 작업을 성공하였습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
									FormItem_Clear();
								}
								else
								{
									BubbleEvent = false;
									return;
								}
							}
						}
					}
					else if (pVal.ItemUID == "Btn01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (HeaderSpaceLineDel(pVal.ItemUID) == false)
							{
								BubbleEvent = false;
								return;
							}
							else
							{
								Create_Form_Code();
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Btn02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1282");
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "WhsCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItmBsort", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItmMsort", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Unit1", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "BaseCode", "");
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "CpNaming")
					{
						if (oForm.Items.Item("CpNaming").Specific.Value == "G")
						{
							oForm.Items.Item("ItmBsort").Specific.Value = "105";
							oForm.Items.Item("ItmMsort").Specific.Value = "10504";
						}
						else if (oForm.Items.Item("CpNaming").Specific.Value == "M")
						{
							oForm.Items.Item("ItmBsort").Specific.Value = "105";
							oForm.Items.Item("ItmMsort").Specific.Value = "10501";
						}
						else if (oForm.Items.Item("CpNaming").Specific.Value == "P")
						{
							oForm.Items.Item("ItmBsort").Specific.Value = "105";
							oForm.Items.Item("ItmMsort").Specific.Value = "10502";
						}
						else if (oForm.Items.Item("CpNaming").Specific.Value == "T")
						{
							oForm.Items.Item("ItmBsort").Specific.Value = "105";
							oForm.Items.Item("ItmMsort").Specific.Value = "10503";
						}
						else if (oForm.Items.Item("CpNaming").Specific.Value == "J")
						{
							oForm.Items.Item("ItmBsort").Specific.Value = "106";
							oForm.Items.Item("ItmMsort").Specific.Value = "10601";
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int loopCount;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					// 고객
					if (pVal.ItemUID == "CardCode" && pVal.ItemChanged == true)
					{
						sQry = "Select CardName From [OCRD] Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();

						if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
						{
							for (loopCount = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; loopCount >= 0; loopCount += -1)
							{
								oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(loopCount, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "선택");
						sQry = " SELECT      U_Minor AS [Code],";
						sQry += "             U_CdName As [Name]";
						sQry += " FROM        [@PS_SY001L]";
						sQry += " WHERE       Code = 'I001'";
						sQry += "             AND U_UseYN = 'Y'";
						sQry += "             AND U_RelCd = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
						sQry += " ORDER BY    U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
						oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
					}

					// 기본창고
					if (pVal.ItemUID == "WhsCode" && pVal.ItemChanged == true)
					{
						sQry = "Select WhsName From [OWHS] Where WhsCode = '" + oForm.Items.Item("WhsCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("WhsName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					}

					// 품목대분류
					if (pVal.ItemUID == "ItmBsort" & pVal.ItemChanged == true)
					{
						sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" + oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("ItmBname").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					}

					// 품목중분류
					if (pVal.ItemUID == "ItmMsort" & pVal.ItemChanged == true)
					{
						sQry = "Select U_CodeName From [@PSH_ITMMSORT] Where U_Code = '" + oForm.Items.Item("ItmMsort").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("ItmMname").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					}

					//기준작번
					if (pVal.ItemUID == "BaseCode" & pVal.ItemChanged == true)
					{
						sQry = "Select ItemName From [OITM] Where ItemCode = '" + oForm.Items.Item("BaseCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("BaseName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// Raise_EVENT_FORM_UNLOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
