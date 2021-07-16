using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 외주제작 작업지시 일괄 등록
	/// </summary>
	internal class PS_PP034 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP034L;
		private int oLast_ColRow;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP034.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP034_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP034");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				SetItems();
				SetItemsData();
				AddMatrixRow(0, true, "");
				SetComboBox();
				EnableMenu();
				EnableFormItem();
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
		/// SetItems
		/// </summary>
		private void SetItems()
		{
			try
			{
				oDS_PS_PP034L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//담당자(사번)
				oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

				//담당자(성명)
				oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// SetItemsData
		/// </summary>
		private void SetItemsData()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		/// <param name="ItemUID"></param>
		private void AddMatrixRow(int oRow, bool RowIserted, string ItemUID)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP034L.InsertRecord(oRow);
				}

				oMat.AddRow();
				oDS_PS_PP034L.Offset = oRow;
				oDS_PS_PP034L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// SetComboBox
		/// </summary>
		private void SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//작업구분(라인)
				sQry = " SELECT      Code,";
				sQry += "             Name";
				sQry += " FROM        [@PSH_ITMBSORT]";
				sQry += " WHERE       U_PudYN = 'Y'";
				sQry += " ORDER BY    Code";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("OrdGbn"), sQry, "", "");

				//외주사유코드
				sQry = "  SELECT     T1.U_Minor AS [Code],";
				sQry += "             T1.U_CdName AS [Value]";
				sQry += "  FROM       [@PS_SY001H] AS T0";
				sQry += "             INNER JOIN";
				sQry += "             [@PS_SY001L] AS T1";
				sQry += "                 ON T0.Code = T1.Code";
				sQry += "  WHERE      T0.Code = 'P201'";
				sQry += "             AND T1.U_UseYN = 'Y'";
				sQry += "  ORDER BY   T1.U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("OutCode"), sQry, "", "");
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
		/// EnableMenu
		/// </summary>
		private void EnableMenu()
		{
			try
			{
				oForm.EnableMenu("1293", true);	//행삭제
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// EnableFormItem
		/// </summary>
		private void EnableFormItem()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.EnableMenu("1281", true);	 //찾기
					oForm.EnableMenu("1282", false); //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);	 //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.EnableMenu("1282", true); //추가
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
		/// FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				if (oUID == "CntcCode")
				{
					sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);
					oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();

				}
				else if (oUID == "Mat01")
				{
					if (oCol == "OrdNum")
					{
						oMat.FlushToDataSource();

						if (CheckDuplicationMatrixData(oDS_PS_PP034L.GetValue("U_ColReg01", oRow - 1).ToString().Trim(), "00", "000") == false)
						{
							errMessage = "1"; //아무것도 안함
							throw new Exception();
						}

						sQry = "PS_PP034_01 '";
						sQry += oDS_PS_PP034L.GetValue("U_ColReg01", oRow - 1).ToString().Trim() +"'";
						oRecordSet.DoQuery(sQry);

						oDS_PS_PP034L.SetValue("U_ColReg01", oRow - 1, oDS_PS_PP034L.GetValue("U_ColReg01", oRow - 1).ToString().Trim()); //작번
						oDS_PS_PP034L.SetValue("U_ColReg02", oRow - 1, "00");	//서브작번1
						oDS_PS_PP034L.SetValue("U_ColReg03", oRow - 1, "000");	//서브작번2
						oDS_PS_PP034L.SetValue("U_ColReg04", oRow - 1, oRecordSet.Fields.Item("BaseType").Value.ToString().Trim()); //기준문서구분
						oDS_PS_PP034L.SetValue("U_ColReg05", oRow - 1, oRecordSet.Fields.Item("BaseNum").Value.ToString().Trim());	//기준문서번호
						oDS_PS_PP034L.SetValue("U_ColReg06", oRow - 1, oRecordSet.Fields.Item("OrdGbn").Value.ToString().Trim());   //작업구분
						oDS_PS_PP034L.SetValue("U_ColDt01", oRow - 1, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //지시일자
						oDS_PS_PP034L.SetValue("U_ColDt02", oRow - 1, Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //지시일자//완료일자
						oDS_PS_PP034L.SetValue("U_ColDt03", oRow - 1, Convert.ToDateTime(oRecordSet.Fields.Item("PP020Dt").Value.ToString().Trim()).ToString("yyyyMMdd")); //작번등록일자
						oDS_PS_PP034L.SetValue("U_ColReg09", oRow - 1, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());	//제품코드
						oDS_PS_PP034L.SetValue("U_ColReg10", oRow - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());	//제품명
						oDS_PS_PP034L.SetValue("U_ColReg11", oRow - 1, oRecordSet.Fields.Item("SjNum").Value.ToString().Trim());	//수주번호
						oDS_PS_PP034L.SetValue("U_ColReg12", oRow - 1, oRecordSet.Fields.Item("SjLine").Value.ToString().Trim());	//수주라인
						oDS_PS_PP034L.SetValue("U_ColReg13", oRow - 1, oRecordSet.Fields.Item("JakMyung").Value.ToString().Trim());	//작번이름
						oDS_PS_PP034L.SetValue("U_ColQty01", oRow - 1, oRecordSet.Fields.Item("ReqWt").Value.ToString().Trim());	//요청수량(중량)
						oDS_PS_PP034L.SetValue("U_ColQty02", oRow - 1, oRecordSet.Fields.Item("SelWt").Value.ToString().Trim());	//지시수량(중량)
						oDS_PS_PP034L.SetValue("U_ColReg16", oRow - 1, oRecordSet.Fields.Item("Comments").Value.ToString().Trim());	//특이사항
						oDS_PS_PP034L.SetValue("U_ColReg17", oRow - 1, oRecordSet.Fields.Item("JakSize").Value.ToString().Trim());	//규격
						oDS_PS_PP034L.SetValue("U_ColReg18", oRow - 1, oRecordSet.Fields.Item("JakUnit").Value.ToString().Trim());	//단위
						oDS_PS_PP034L.SetValue("U_ColQty03", oRow - 1, oRecordSet.Fields.Item("ReqQty").Value.ToString().Trim());	//구매청구수량
						oDS_PS_PP034L.SetValue("U_ColReg19", oRow - 1, oRecordSet.Fields.Item("ReqUnit").Value.ToString().Trim());  //구매청구단위
						oDS_PS_PP034L.SetValue("U_ColDt04", oRow - 1, Convert.ToDateTime(oRecordSet.Fields.Item("ReqDt").Value.ToString().Trim()).ToString("yyyyMMdd")); //구매청구납기

						if (oMat.RowCount == oRow && !string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColReg01", oRow - 1).ToString().Trim()))
						{
							AddMatrixRow(oRow, false, "");
						}

						oMat.LoadFromDataSource();
						oMat.Columns.Item("OrdNum").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					}
					else if (oCol == "OrdSub1" || oCol == "OrdSub2")
					{

						oMat.FlushToDataSource();

						sQry = "PS_PP034_02 '";
						sQry += oDS_PS_PP034L.GetValue("U_ColReg01", oRow - 1).ToString().Trim() +"','";
						sQry += oDS_PS_PP034L.GetValue("U_ColReg02", oRow - 1).ToString().Trim() +"','";
						sQry += oDS_PS_PP034L.GetValue("U_ColReg03", oRow - 1).ToString().Trim() +"'";
						oRecordSet.DoQuery(sQry);

						oDS_PS_PP034L.SetValue("U_ColReg02", oRow - 1, oDS_PS_PP034L.GetValue("U_ColReg02", oRow - 1).ToString().Trim()); //서브작번1
						oDS_PS_PP034L.SetValue("U_ColReg03", oRow - 1, oDS_PS_PP034L.GetValue("U_ColReg03", oRow - 1).ToString().Trim()); //서브작번2
						oDS_PS_PP034L.SetValue("U_ColReg04", oRow - 1, oRecordSet.Fields.Item("BaseType").Value.ToString().Trim()); //기준문서구분
						oDS_PS_PP034L.SetValue("U_ColReg05", oRow - 1, oRecordSet.Fields.Item("BaseNum").Value.ToString().Trim());	//기준문서번호
						oDS_PS_PP034L.SetValue("U_ColReg06", oRow - 1, oRecordSet.Fields.Item("OrdGbn").Value.ToString().Trim());   //작업구분
						oDS_PS_PP034L.SetValue("U_ColDt01", oRow - 1, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //지시일자
						oDS_PS_PP034L.SetValue("U_ColDt02", oRow - 1, Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //완료일자
						oDS_PS_PP034L.SetValue("U_ColDt03", oRow - 1, Convert.ToDateTime(oRecordSet.Fields.Item("PP020Dt").Value.ToString().Trim()).ToString("yyyyMMdd")); //작번등록일자
						oDS_PS_PP034L.SetValue("U_ColReg09", oRow - 1, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim()); //제품코드
						oDS_PS_PP034L.SetValue("U_ColReg10", oRow - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //제품명
						oDS_PS_PP034L.SetValue("U_ColReg11", oRow - 1, oRecordSet.Fields.Item("SjNum").Value.ToString().Trim());	//수주번호
						oDS_PS_PP034L.SetValue("U_ColReg12", oRow - 1, oRecordSet.Fields.Item("SjLine").Value.ToString().Trim());	//수주라인
						oDS_PS_PP034L.SetValue("U_ColReg13", oRow - 1, oRecordSet.Fields.Item("JakMyung").Value.ToString().Trim()); //작번이름
						oDS_PS_PP034L.SetValue("U_ColQty01", oRow - 1, oRecordSet.Fields.Item("ReqWt").Value.ToString().Trim());	//요청수량(중량)
						oDS_PS_PP034L.SetValue("U_ColQty02", oRow - 1, oRecordSet.Fields.Item("SelWt").Value.ToString().Trim());	//지시수량(중량)
						oDS_PS_PP034L.SetValue("U_ColReg16", oRow - 1, oRecordSet.Fields.Item("Comments").Value.ToString().Trim()); //특이사항
						oDS_PS_PP034L.SetValue("U_ColReg17", oRow - 1, oRecordSet.Fields.Item("JakSize").Value.ToString().Trim());	//규격
						oDS_PS_PP034L.SetValue("U_ColReg18", oRow - 1, oRecordSet.Fields.Item("JakUnit").Value.ToString().Trim());	//단위
						oDS_PS_PP034L.SetValue("U_ColQty03", oRow - 1, oRecordSet.Fields.Item("ReqQty").Value.ToString().Trim());	//구매청구수량
						oDS_PS_PP034L.SetValue("U_ColReg19", oRow - 1, oRecordSet.Fields.Item("ReqUnit").Value.ToString().Trim());  //구매청구단위
						oDS_PS_PP034L.SetValue("U_ColDt04", oRow - 1, Convert.ToDateTime(oRecordSet.Fields.Item("ReqDt").Value.ToString().Trim()).ToString("yyyyMMdd")); //구매청구납기

						oMat.LoadFromDataSource();

						if (oCol == "OrdSub1")
						{
							oMat.Columns.Item("OrdSub1").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else if (oCol == "OrdSub2")
						{
							oMat.Columns.Item("OrdSub2").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
					}

					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				if (errMessage == "1")
				{
					//PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// CheckHeadData
		/// </summary>
		/// <returns></returns>
		private bool CheckHeadData()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
				{
					errMessage = "사업장은 필수입니다. 확인하세요.";
					throw new Exception();
				}

				if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "담당자는 필수입니다. 확인하세요.";
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
		/// CheckLineData
		/// </summary>
		/// <returns></returns>
		private bool CheckLineData()
		{
			bool functionReturnValue = false;

			int loopCount;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 2; loopCount++)
				{
					if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColReg01", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 작번이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColReg02", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 서브작번1이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColReg03", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 서브작번2가 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColDt01", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 지시일자가 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColDt02", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 완료일자가 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColQty02", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 지시수량(중량)이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColQty03", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 구매청수량이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColReg19", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 구매청구단위가 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColDt04", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 구매청구납기가 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColReg20", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 외주사유코드가 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oDS_PS_PP034L.GetValue("U_ColReg21", loopCount).ToString().Trim()))
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 외주사유내용이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else if (CheckValidItemCode(oDS_PS_PP034L.GetValue("U_ColReg09", loopCount).ToString().Trim()) == "UnAuthorized")
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의  품목은 미승인 품목입니다. 승인 후 구매요청을 진행하십시오.";
						throw new Exception();
					}
					else if (CheckValidItemCode(oDS_PS_PP034L.GetValue("U_ColReg09", loopCount).ToString().Trim()) == "UnUsed")
					{
						errMessage = Convert.ToString(loopCount + 1) + "번 라인의 품목은 비활성 품목입니다. 확인하세요.";
						throw new Exception();
					}
				}
				oMat.LoadFromDataSource();

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
		/// CheckDuplicationDBData
		/// </summary>
		/// <returns></returns>
		private bool CheckDuplicationDBData()
		{
			bool functionReturnValue = false;

			short loopCount;
			string sQry;
			string OrdNum;
			string OrdSub1;
			string OrdSub2;
			string errMessage = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 2; loopCount++)
				{
					OrdNum = oDS_PS_PP034L.GetValue("U_ColReg01", loopCount).ToString().Trim();
					OrdSub1 = oDS_PS_PP034L.GetValue("U_ColReg02", loopCount).ToString().Trim();
					OrdSub2 = oDS_PS_PP034L.GetValue("U_ColReg03", loopCount).ToString().Trim();

					sQry = " SELECT    COUNT(*)";
					sQry += " FROM      [@PS_PP030H] AS T0";
					sQry += " WHERE     T0.U_OrdNum = '" + OrdNum + "'";
					sQry += "           AND T0.U_OrdSub1 = '" + OrdSub1 + "'";
					sQry += "           AND T0.U_OrdSub2 = '" + OrdSub2 + "'";
					oRecordSet.DoQuery(sQry);

					if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) > 0)
					{
						errMessage = Convert.ToString(loopCount + 1) + "행은 이미 작업지시에 등록된 작번입니다. 확인하십시오.";
						throw new Exception();
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// CheckDuplicationMatrixData
		/// </summary>
		/// <param name="pOrdNum"></param>
		/// <param name="pOrdSub1"></param>
		/// <param name="pOrdSub2"></param>
		/// <returns></returns>
		private bool CheckDuplicationMatrixData(string pOrdNum, string pOrdSub1, string pOrdSub2)
		{
			bool functionReturnValue = false;

			int loopCount;
			string OrdNum;
			string OrdSub1;
			string OrdSub2;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 2; loopCount++)
				{
					OrdNum = oDS_PS_PP034L.GetValue("U_ColReg01", loopCount).ToString().Trim();
					OrdSub1 = oDS_PS_PP034L.GetValue("U_ColReg02", loopCount).ToString().Trim();
					OrdSub2 = oDS_PS_PP034L.GetValue("U_ColReg03", loopCount).ToString().Trim();

					if (pOrdNum == OrdNum && pOrdSub1 == OrdSub1 && pOrdSub2 == OrdSub2)
					{
						errMessage = "작번 [" + OrdNum + "-" + OrdSub1 + "-" + OrdSub2 + "]은(는) 이미 Matrix에 추가된 작번입니다. 확인하십시오.";
						throw new Exception();
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
		/// CheckDocDate
		/// 선행프로세스와 일자 비교
		/// </summary>
		/// <returns></returns>
		private bool CheckDocDate()
		{
			bool functionReturnValue = false;

			string sQry;
			int loopCount;
			string BaseEntry;
			string BaseLine;
			string DocType;
			string CurDocDate;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BaseLine = "";
				DocType = "PS_PP030";
				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 2; loopCount++)
				{

					BaseEntry = oDS_PS_PP034L.GetValue("U_ColReg05", loopCount).ToString().Trim();
					CurDocDate = oDS_PS_PP034L.GetValue("U_ColDt01", loopCount).ToString().Trim();

					sQry = " EXEC PS_Z_CHECK_DATE '";
					sQry += BaseEntry + "','";
					sQry += BaseLine + "','";
					sQry += DocType + "','";
					sQry += CurDocDate + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.Fields.Item("ReturnValue").Value.ToString().Trim() == "False")
					{
						errMessage = Convert.ToString(loopCount + 1) + "행의 지시일자가 작번등록일보다 빠릅니다. 확인하십시오.";
						throw new Exception();
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
		/// CheckValidItemCode
		/// </summary>
		/// <param name="pItemCode"></param>
		/// <returns></returns>
		private string CheckValidItemCode(string pItemCode)
		{
			string functionReturnValue = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = " SELECT  frozenTo,";
				sQry += "         frozenFor";
				sQry += " FROM    OITM";
				sQry += " WHERE   ItemCode = '" + pItemCode + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.Fields.Item("frozenTo").Value.ToString().Trim() == "2999-12-31" && oRecordSet.Fields.Item("frozenFor").Value.ToString().Trim() == "Y")
				{
					functionReturnValue = "UnAuthorized"; //미승인
				}
				else if (oRecordSet.Fields.Item("frozenTo").Value.ToString().Trim() == "2899-12-31" && oRecordSet.Fields.Item("frozenFor").Value.ToString().Trim() == "Y")
				{
					functionReturnValue = "UnUsed"; //미사용
				}
				else
				{
					functionReturnValue = "Authorized"; //승인
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
			return functionReturnValue;
		}

		/// <summary>
		/// AddData
		/// 데이터 INSERT
		/// </summary>
		private void AddData()
		{
			int loopCount;
			string sQry;

			string OrdNum;   //작번
			string OrdSub1;  //서브작번1
			string OrdSub2;  //서브작번2
			string BaseType; //기준문서구분
			string baseNum;  //기준문서번호
			string OrdGbn;   //작업구분
			string DocDate;  //지시일자
			string DueDate;  //완료일자
			string ItemCode; //제품코드
			string ItemName; //제품명
			string sjNum;    //수주번호
			string sjLine;   //수주라인
			string JakMyung; //작번이름
			double reqWt;    //요청수량(중량)
			double selWt;    //지시수량(중량)
			string Comments; //특이사항
			string JakSize;  //규격
			string JakUnit;  //단위
			string BPLId;    //사업장코드
			string UserSign; //UserSign
			string UserID;   //UserID
			string CntcCode; //사번
			string CntcName; //성명

			//구매요청 등록 변수
			double ReqQty;  //구매청구수량
			string reqUnit; //구매청구단위
			string reqDt;   //구매청구납기
			string outCode; //외주사유코드
			string outNote; //외주사유내용
			string reqNote; //구매청구비고

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
				BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();

				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 2; loopCount++)
				{

					OrdNum = oDS_PS_PP034L.GetValue("U_ColReg01", loopCount).ToString().Trim();
					OrdSub1 = oDS_PS_PP034L.GetValue("U_ColReg02", loopCount).ToString().Trim();
					OrdSub2 = oDS_PS_PP034L.GetValue("U_ColReg03", loopCount).ToString().Trim();
					BaseType = oDS_PS_PP034L.GetValue("U_ColReg04", loopCount).ToString().Trim();
					baseNum = oDS_PS_PP034L.GetValue("U_ColReg05", loopCount).ToString().Trim();
					OrdGbn = oDS_PS_PP034L.GetValue("U_ColReg06", loopCount).ToString().Trim();
					DocDate = oDS_PS_PP034L.GetValue("U_ColDt01", loopCount).ToString().Trim();
					DueDate = oDS_PS_PP034L.GetValue("U_ColDt02", loopCount).ToString().Trim();
					ItemCode = oDS_PS_PP034L.GetValue("U_ColReg09", loopCount).ToString().Trim();
					ItemName = oDS_PS_PP034L.GetValue("U_ColReg10", loopCount).ToString().Trim();
					sjNum = oDS_PS_PP034L.GetValue("U_ColReg11", loopCount).ToString().Trim();
					sjLine = oDS_PS_PP034L.GetValue("U_ColReg12", loopCount).ToString().Trim();
					JakMyung = oDS_PS_PP034L.GetValue("U_ColReg13", loopCount).ToString().Trim();
					reqWt = Convert.ToDouble(oDS_PS_PP034L.GetValue("U_ColQty01", loopCount).ToString().Trim());
					selWt = Convert.ToDouble(oDS_PS_PP034L.GetValue("U_ColQty02", loopCount).ToString().Trim());
					Comments = oDS_PS_PP034L.GetValue("U_ColReg16", loopCount).ToString().Trim();
					JakSize = oDS_PS_PP034L.GetValue("U_ColReg17", loopCount).ToString().Trim();
					JakUnit = oDS_PS_PP034L.GetValue("U_ColReg18", loopCount).ToString().Trim();
					UserSign = Convert.ToString(PSH_Globals.oCompany.UserSignature);
					UserID = PSH_Globals.SBO_Application.Company.UserName;

					ReqQty = Convert.ToDouble(oDS_PS_PP034L.GetValue("U_ColQty03", loopCount).ToString().Trim());
					reqUnit = oDS_PS_PP034L.GetValue("U_ColReg19", loopCount).ToString().Trim();
					reqDt = oDS_PS_PP034L.GetValue("U_ColDt04", loopCount).ToString().Trim();
					outCode = oDS_PS_PP034L.GetValue("U_ColReg20", loopCount).ToString().Trim();
					outNote = oDS_PS_PP034L.GetValue("U_ColReg21", loopCount).ToString().Trim();
					reqNote = oDS_PS_PP034L.GetValue("U_ColReg22", loopCount).ToString().Trim();

					sQry = " EXEC PS_PP034_03 '";
					sQry += OrdNum + "','";
					sQry += OrdSub1 + "','";
					sQry += OrdSub2 + "','";
					sQry += BaseType + "','";
					sQry += baseNum + "','";
					sQry += OrdGbn + "','";
					sQry += DocDate + "','";
					sQry += DueDate + "','";
					sQry += ItemCode + "','";
					sQry += ItemName + "','";
					sQry += sjNum + "','";
					sQry += sjLine + "','";
					sQry += JakMyung + "',";
					sQry += reqWt + ",";
					sQry += selWt + ",'";
					sQry += Comments + "','";
					sQry += JakSize + "','";
					sQry += JakUnit + "','";
					sQry += BPLId + "','";
					sQry += CntcCode + "','";
					sQry += CntcName + "','";
					sQry += UserSign + "','";
					sQry += UserID + "',";
					sQry += ReqQty + ",'";
					sQry += reqUnit + "','";
					sQry += reqDt + "','";
					sQry += outCode + "','";
					sQry += outNote + "','";
					sQry += reqNote + "'";
					oRecordSet.DoQuery(sQry); //DB저장

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToString(oMat.VisualRowCount - 1) + "건 저장중...!";
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("처리 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
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
						if (CheckHeadData() == false)
						{
							BubbleEvent = false;
							return;
						}

						if (CheckLineData() == false)
						{
							BubbleEvent = false;
							return;
						}

						if (CheckDocDate() == false)
						{
							BubbleEvent = false;
							return;
						}

						if (CheckDuplicationDBData() == false)
						{
							BubbleEvent = false;
							return;
						}
						AddData(); //필수 입력 조건 모두 만족하면 데이터 입력
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "BtnAdd")
					{
						oMat.Clear(); //추가 후 Matrix 초기화
						oMat.FlushToDataSource();
						AddMatrixRow(0, true, "");
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
					if (pVal.ItemUID == "Mat01")
					{
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrdNum");
					}
					else
					{
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
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
							oLast_ColRow = pVal.Row;
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
					}
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP034L);
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
						case "1285": //복원
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
						case "1281": //찾기
							EnableFormItem();
							break;
						case "1282": //추가
							EnableFormItem();
							AddMatrixRow(0, true, "");
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "1293": //행삭제
							if (oMat.RowCount != oMat.VisualRowCount)
							{
								oMat.FlushToDataSource();
								for (int loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
								{
									oDS_PS_PP034L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
								}
								oDS_PS_PP034L.RemoveRecord(oDS_PS_PP034L.Size - 1);
								oMat.Clear();
								oMat.LoadFromDataSource();
							}
							break;
						case "7169": //엑셀 내보내기
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
