using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 장비별 목표금액 수정관리(생산)
	/// </summary>
	internal class PS_PP562 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP562B; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP562.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP562_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP562");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP562_CreateItems();
				PS_PP562_SetComboBox();
				PS_PP562_ResizeForm();
				PS_PP562_LoadCaption();

				oForm.EnableMenu("1283", false);	// 삭제
				oForm.EnableMenu("1286", false);	// 닫기
				oForm.EnableMenu("1287", false);	// 복제
				oForm.EnableMenu("1285", false);	// 복원
				oForm.EnableMenu("1284", false);	// 취소
				oForm.EnableMenu("1293", false);	// 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);

				PS_PP562_ResetForm();  //폼초기화

				oForm.Items.Item("BaseEntry").Visible = false; 	//BaseEntry 비활성
				oForm.Items.Item("BaseLine").Visible = false;	//BaseLine 비활성
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
		/// PS_PP562_CreateItems
		/// </summary>
		private void PS_PP562_CreateItems()
		{
			try
			{
				oDS_PS_PP562B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				// 메트릭스 개체 할당
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//등록정보
				//문서번호
				oForm.DataSources.UserDataSources.Add("BaseDL", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("BaseDL").Specific.DataBind.SetBound(true, "", "BaseDL");

				//비용구분
				oForm.DataSources.UserDataSources.Add("AmtCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("AmtCls").Specific.DataBind.SetBound(true, "", "AmtCls");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum").Specific.DataBind.SetBound(true, "", "OrdNum");

				//서브작번1
				oForm.DataSources.UserDataSources.Add("OrdSub1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("OrdSub1").Specific.DataBind.SetBound(true, "", "OrdSub1");

				//서브작번2
				oForm.DataSources.UserDataSources.Add("OrdSub2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("OrdSub2").Specific.DataBind.SetBound(true, "", "OrdSub2");

				//작지명
				oForm.DataSources.UserDataSources.Add("OrdName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("OrdName").Specific.DataBind.SetBound(true, "", "OrdName");

				//목표금액
				oForm.DataSources.UserDataSources.Add("Amount", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("Amount").Specific.DataBind.SetBound(true, "", "Amount");

				//비고
				oForm.DataSources.UserDataSources.Add("Comment", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("Comment").Specific.DataBind.SetBound(true, "", "Comment");

				//기준문서번호
				oForm.DataSources.UserDataSources.Add("BaseEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BaseEntry").Specific.DataBind.SetBound(true, "", "BaseEntry");

				//기준문서라인번호
				oForm.DataSources.UserDataSources.Add("BaseLine", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BaseLine").Specific.DataBind.SetBound(true, "", "BaseLine");

				//조회정보
				//비용구분
				oForm.DataSources.UserDataSources.Add("SAmtCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SAmtCls").Specific.DataBind.SetBound(true, "", "SAmtCls");

				//작번
				oForm.DataSources.UserDataSources.Add("SOrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SOrdNum").Specific.DataBind.SetBound(true, "", "SOrdNum");

				//서브작번1
				oForm.DataSources.UserDataSources.Add("SOrdSub1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2);
				oForm.Items.Item("SOrdSub1").Specific.DataBind.SetBound(true, "", "SOrdSub1");

				//서브작번2
				oForm.DataSources.UserDataSources.Add("SOrdSub2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 3);
				oForm.Items.Item("SOrdSub2").Specific.DataBind.SetBound(true, "", "SOrdSub2");

				//작지명
				oForm.DataSources.UserDataSources.Add("SOrdName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("SOrdName").Specific.DataBind.SetBound(true, "", "SOrdName");

				//비고
				oForm.DataSources.UserDataSources.Add("SComment", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("SComment").Specific.DataBind.SetBound(true, "", "SComment");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP562_SetComboBox
		/// </summary>
		private void PS_PP562_SetComboBox()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//등록정보
				//비용구분
				oForm.Items.Item("AmtCls").Specific.ValidValues.Add("%", "선택");
				sQry = " SELECT      T0.U_Minor, ";
				sQry += "             T0.U_CdName";
				sQry += " FROM        [@PS_SY001L] T0";
				sQry += " WHERE       T0.Code = 'P210'";
				sQry += "             AND T0.U_UseYN = 'Y'";
				sQry += " ORDER BY    T0.U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("AmtCls").Specific, sQry, "", false, false);
				oForm.Items.Item("AmtCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//조회정보
				//비용구분
				oForm.Items.Item("SAmtCls").Specific.ValidValues.Add("%", "전체");
				sQry = " SELECT      T0.U_Minor, ";
				sQry += "             T0.U_CdName";
				sQry += " FROM        [@PS_SY001L] T0";
				sQry += " WHERE       T0.Code = 'P210'";
				sQry += "             AND T0.U_UseYN = 'Y'";
				sQry += " ORDER BY    T0.U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("SAmtCls").Specific, sQry, "", false, false);
				oForm.Items.Item("SAmtCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP562_ResizeForm
		/// </summary>
		private void PS_PP562_ResizeForm()
		{
			try
			{
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP562_LoadCaption
		/// </summary>
		private void PS_PP562_LoadCaption()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
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
		/// PS_PP562_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP562_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP562B.InsertRecord(oRow);
				}

				oMat.AddRow();
				oDS_PS_PP562B.Offset = oRow;
				oDS_PS_PP562B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP562_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP562_DelHeaderSpaceLine()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (oForm.Items.Item("AmtCls").Specific.Value.ToString().Trim() == "%") //비용구분
				{
					errMessage = "비용구분은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Items.Item("AmtCls").Specific.Value.ToString().Trim() == "01" && oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim() == "") //일반 & 작번
				{
					errMessage = "작번은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Items.Item("Amount").Specific.Value.ToString().Trim() == "%") //목표금액
				{
					errMessage = "금액은 필수사항입니다. 확인하세요.";
					throw new Exception();
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
		/// PS_PP562_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP562_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "BaseDL":
						sQry = " SELECT      T0.U_AmtCls AS [AmtCls],";
						sQry += "             T0.U_OrdNum AS [OrdNum],";
						sQry += "             T0.U_OrdSub1 AS [OrdSub1],";
						sQry += "             T0.U_OrdSub2 AS [OrdSub2],";
						sQry += "             T0.U_OrdName AS [OrdName],";
						sQry += "             T0.U_TrgtAmt AS [Amount],";
						sQry += "             T0.U_Comment AS [Comment],";
						sQry += "             T0.DocEntry AS [BaseEntry],";
						sQry += "             T0.U_LineUID AS [BaseLine]";
						sQry += " FROM        [@PS_PP560L] AS T0";
						sQry += " WHERE       CONVERT(VARCHAR(20), T0.DocEntry) + '-' + CONVERT(VARCHAR(20), T0.U_LineUID) = '" + oForm.DataSources.UserDataSources.Item("BaseDL").Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("AmtCls").Enabled = true;
						oForm.DataSources.UserDataSources.Item("AmtCls").Value = oRecordSet.Fields.Item("AmtCls").Value.ToString().Trim();		 //비용구분
						oForm.Items.Item("AmtCls").Enabled = false;																				 
						oForm.DataSources.UserDataSources.Item("OrdNum").Value = oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim();       //작번
						oForm.DataSources.UserDataSources.Item("OrdSub1").Value = oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim();	 //서브작번1
						oForm.DataSources.UserDataSources.Item("OrdSub2").Value = oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim();	 //서브작번2
						oForm.DataSources.UserDataSources.Item("OrdName").Value = oRecordSet.Fields.Item("OrdName").Value.ToString().Trim();	 //작명
						oForm.DataSources.UserDataSources.Item("Amount").Value = oRecordSet.Fields.Item("Amount").Value.ToString().Trim();		 //목표금액
						oForm.DataSources.UserDataSources.Item("Comment").Value = oRecordSet.Fields.Item("Comment").Value.ToString().Trim();	 //비고
						oForm.DataSources.UserDataSources.Item("BaseEntry").Value = oRecordSet.Fields.Item("BaseEntry").Value.ToString().Trim(); //BaseEntry
						oForm.DataSources.UserDataSources.Item("BaseLine").Value = oRecordSet.Fields.Item("BaseLine").Value.ToString().Trim();	 //BaseLine
						break;
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
		/// PS_PP562_ResetForm
		/// </summary>
		private void PS_PP562_ResetForm()
		{
			try
			{
				oForm.Freeze(true);

				//기준정보
				oForm.DataSources.UserDataSources.Item("BaseDL").Value = "";	//문서번호
				oForm.DataSources.UserDataSources.Item("AmtCls").Value = "%";	//비용구분
				oForm.DataSources.UserDataSources.Item("OrdNum").Value = "";	//작번
				oForm.DataSources.UserDataSources.Item("OrdSub1").Value = "";	//서브작번1
				oForm.DataSources.UserDataSources.Item("OrdSub2").Value = "";	//서브작번2
				oForm.DataSources.UserDataSources.Item("OrdName").Value = "";	//작지명
				oForm.DataSources.UserDataSources.Item("Amount").Value = Convert.ToString(0); //목표금액
				oForm.DataSources.UserDataSources.Item("Comment").Value = "";	//비고
				oForm.DataSources.UserDataSources.Item("BaseEntry").Value = "";	//기준문서번호
				oForm.DataSources.UserDataSources.Item("BaseLine").Value = "";	//기준문서라인번호

				//조회정보
				oForm.DataSources.UserDataSources.Item("SAmtCls").Value = "%";  //비용구분
				oForm.DataSources.UserDataSources.Item("SOrdNum").Value = "";   //작번
				oForm.DataSources.UserDataSources.Item("SOrdSub1").Value = "";  //서브작번1
				oForm.DataSources.UserDataSources.Item("SOrdSub2").Value = "";  //서브작번2
				oForm.DataSources.UserDataSources.Item("SOrdName").Value = "";  //작지명
				oForm.DataSources.UserDataSources.Item("SComment").Value = "";  //비고(%)

				oForm.Items.Item("BaseDL").Click();
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
		/// PS_PP562_DeleteData  기본정보 삭제(Delete 사용 안함)
		/// </summary>
		private void PS_PP562_DeleteData()
		{
			string sQry;
			string errMessage = string.Empty;
			string DocEntry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

					sQry = "SELECT COUNT(*) FROM [Z_PS_PP560_01] WHERE DocEntry = '" + DocEntry + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						errMessage = "삭제대상이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else
					{
						sQry = "EXEC PS_PP562_04 '" + DocEntry + "'";
						oRecordSet.DoQuery(sQry);
					}
				}

				PSH_Globals.SBO_Application.MessageBox("삭제 완료!");
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP562_UpdateData  기본정보를 수정(Update 사용 안함)
		/// </summary>
		/// <returns></returns>
		private bool PS_PP562_UpdateData()
		{
			bool functionReturnValue = false;
			string sQry = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.MessageBox("수정 완료!");

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}

			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP562_AddData 데이터 INSERT
		/// </summary>
		/// <returns></returns>
		private bool PS_PP562_AddData()
		{
			bool functionReturnValue = false;
			string sQry;			
			string AmtCls;	//비용구분
			string OrdNum;	//작번
			string OrdSub1;	//서브작번1
			string OrdSub2;	//서브작번2
			string OrdName;	//작지명
			decimal Amount;	//금액
			string Comment;	//비고
			string UpdateUserCD; //등록자 사번
			string DocCls;	//DocCls
			string BaseEntry; //기준문서번호
			string BaseLine; //기준문서라인번호
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				AmtCls = oForm.Items.Item("AmtCls").Specific.Value.ToString().Trim();
				OrdNum  = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				OrdSub1 = oForm.Items.Item("OrdSub1").Specific.Value.ToString().Trim();
				OrdSub2 = oForm.Items.Item("OrdSub2").Specific.Value.ToString().Trim();
				OrdName = oForm.Items.Item("OrdName").Specific.Value.ToString().Trim();
				Amount = Convert.ToDecimal(oForm.Items.Item("Amount").Specific.Value.ToString().Trim());
				Comment = oForm.Items.Item("Comment").Specific.Value.ToString().Trim();
				UpdateUserCD = dataHelpClass.User_MSTCOD();
				DocCls = "PP562";
				BaseEntry = oForm.Items.Item("BaseEntry").Specific.Value.ToString().Trim();
				BaseLine  = oForm.Items.Item("BaseLine").Specific.Value.ToString().Trim();

				sQry = "EXEC [PS_PP562_02] '";
				sQry += AmtCls + "','";
				sQry += OrdNum + "','";
				sQry += OrdSub1 + "','";
				sQry += OrdSub2 + "','";
				sQry += OrdName + "','";
				sQry += Amount + "','";
				sQry += Comment + "','";
				sQry += UpdateUserCD + "','";
				sQry += DocCls + "','";
				sQry += BaseEntry + "','";
				sQry += BaseLine + "'";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.MessageBox("등록 완료!");

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP562_MTX01 데이터 조회
		/// </summary>
		private void PS_PP562_MTX01()
		{
			int i;
			string sQry;
			string errMessage = string.Empty;
			string SAmtCls;	 //비용구분
			string SOrdNum;	 //작번
			string SOrdSub1; //서브작번1
			string SOrdSub2; //서브작번2
			string SComment; //비고(%)
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				SAmtCls = oForm.DataSources.UserDataSources.Item("SAmtCls").Value.ToString().Trim();
				SOrdNum = oForm.DataSources.UserDataSources.Item("SOrdNum").Value.ToString().Trim();
				SOrdSub1 = oForm.DataSources.UserDataSources.Item("SOrdSub1").Value.ToString().Trim();
				SOrdSub2 = oForm.DataSources.UserDataSources.Item("SOrdSub2").Value.ToString().Trim();
				SComment = oForm.DataSources.UserDataSources.Item("SComment").Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				oForm.Freeze(true);

				sQry = "EXEC [PS_PP562_01] '";
				sQry += SAmtCls + "','";
				sQry += SOrdNum + "','";
				sQry += SOrdSub1 + "','";
				sQry += SOrdSub2 + "','";
				sQry += SComment + "'";

				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP562B.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_PP562_LoadCaption();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP562B.Size)
					{
						oDS_PS_PP562B.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP562B.Offset = i;
					oDS_PS_PP562B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP562B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("AmtCls").Value.ToString().Trim());		//비용구분
					oDS_PS_PP562B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());		//작지번호
					oDS_PS_PP562B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());		//서브작번1
					oDS_PS_PP562B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());		//서브작번2
					oDS_PS_PP562B.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("OrdName").Value.ToString().Trim());		//작지명
					oDS_PS_PP562B.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Amount").Value.ToString().Trim());		//목표금액
					oDS_PS_PP562B.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("Comment").Value.ToString().Trim());		//비고
					oDS_PS_PP562B.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("UpdateUser").Value.ToString().Trim());	//수정자
					oDS_PS_PP562B.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("UpdateDate").Value.ToString().Trim());	//수정일
					oDS_PS_PP562B.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("UpdateTime").Value.ToString().Trim());	//수정시간
					oDS_PS_PP562B.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("BaseDL").Value.ToString().Trim());		//기준문서
					oDS_PS_PP562B.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("MainOrd").Value.ToString().Trim());		//메인작번

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
					if (pVal.ItemUID == "BtnAdd")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_PP562_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_PP562_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_PP562_ResetForm();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_PP562_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_PP562_UpdateData() == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_PP562_ResetForm();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

							PS_PP562_LoadCaption();
							PS_PP562_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSearch")
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_PP562_LoadCaption();
						PS_PP562_MTX01();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "BaseDL", "");
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
							oMat.SelectRow(pVal.Row, true, false);
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
					PS_PP562_FlushToItemValue(pVal.ItemUID, 0, "");
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
							PS_PP562_FlushToItemValue(pVal.ItemUID,0, "");
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
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUIDl"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUIDl, SAPbouiCOM.ItemEvent pVal, bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP562_ResizeForm();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP562B);
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
							PS_PP562_ResetForm();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_PP562_LoadCaption();
							break;
						case "1285": //복원
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							oForm.Freeze(true);
							PS_PP562_AddMatrixRow(oMat.VisualRowCount, false);
							oForm.Freeze(false);
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
						case "1285": //복원
							break;
						case "1293": //행삭제
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
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 이후 처리
							oForm.Freeze(true);
							oDS_PS_PP562B.RemoveRecord(oDS_PS_PP562B.Size - 1);
							oMat.LoadFromDataSource();
							oForm.Freeze(false);
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
