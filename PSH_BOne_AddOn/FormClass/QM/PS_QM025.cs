using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 검사성적서 납품처변경
	/// </summary>
	internal class PS_QM025 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_QM025L; //등록라인

		/// <summary>
		/// Form 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM025.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM025_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM025");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM025_CreateItems();
				PS_QM025_ComboBox_Setting();
				PS_QM025_LoadCaption();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
		/// PS_QM025_CreateItems
		/// </summary>
		private void PS_QM025_CreateItems()
		{
			try
			{
				oDS_PS_QM025L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM025_ComboBox_Setting
		/// </summary>
		private void PS_QM025_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_QM025_LoadCaption
		/// </summary>
		private void PS_QM025_LoadCaption()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("Btn01").Specific.Caption = "수정";
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("Btn01").Specific.Caption = "확인";
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM025_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_QM025_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string CardCode;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "Mat01":
						if (oCol == "CardCode")
						{
							CardCode = oMat.Columns.Item("CardCode").Cells.Item(oRow).Specific.Value.ToString().Trim();
							sQry = "Select CardName From OCRD Where CardCode = '" + CardCode + "'";
							oRecordSet.DoQuery(sQry);
							oMat.Columns.Item("CardName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_QM025_LoadData
		/// </summary>
		private void PS_QM025_LoadData()
		{
			int i;
			string PackNoF;
			string BPLId;
			string PackNoT;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				PackNoF = oForm.Items.Item("PackNoF").Specific.Value.ToString().Trim();
				PackNoT = oForm.Items.Item("PackNoT").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_QM025_01] '" + BPLId + "', '" + PackNoF + "', '" + PackNoT + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_QM025L.Clear();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_QM025L.Size)
					{
						oDS_PS_QM025L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_QM025L.Offset = i;
					oDS_PS_QM025L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM025L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("U_OrdNum").Value.ToString().Trim());
					oDS_PS_QM025L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("U_InspDate").Value.ToString().Trim()).ToString("yyyyMMdd"));
					oDS_PS_QM025L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("U_CardCode").Value.ToString().Trim());
					oDS_PS_QM025L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("U_CardName").Value.ToString().Trim());
					oDS_PS_QM025L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim());
					oDS_PS_QM025L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim());
					oRecordSet.MoveNext();

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
		/// PS_QM025_CH_QM020
		/// </summary>
		/// <returns></returns>
		private bool PS_QM025_CH_QM020()
		{
			bool ReturnValue = false;
			int i;
			string PackNoF;
			string BPLId;
			string PackNoT;
			string CardCode;
			string OrdNum;
			string CardName;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				PackNoF = oForm.Items.Item("PackNoF").Specific.Value.ToString().Trim();
				PackNoT = oForm.Items.Item("PackNoT").Specific.Value.ToString().Trim();

				oMat.FlushToDataSource();

				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				PSH_Globals.oCompany.StartTransaction();

				ProgressBar01.Text = "수정중.";

				for (i = 0; i <= oMat.RowCount - 1; i++)
				{
					CardCode = oDS_PS_QM025L.GetValue("U_ColReg02", i).ToString().Trim();
					sQry = "Select CardName From OCRD Where CardCode = '" + CardCode + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						errMessage = "납품처에 잘못된 자료가 있습니다. 확인하세요.";
						throw new Exception();
					}
				}

				for (i = 0; i <= oMat.RowCount - 1; i++)
				{
					OrdNum = oDS_PS_QM025L.GetValue("U_ColReg01", i).ToString().Trim(); //작업지시번호
					CardCode = oDS_PS_QM025L.GetValue("U_ColReg02", i).ToString().Trim(); //납품처코드

					sQry = "Select CardName From OCRD Where CardCode = '" + CardCode + "'";
					oRecordSet.DoQuery(sQry);
					CardName = oRecordSet.Fields.Item(0).Value.ToString().Trim();

					sQry = "Update [@PS_QM610H]";
					sQry += " set U_CardCode = '" + CardCode + "',";
					sQry += " U_CardName = '" + CardName + "'";
					sQry += " Where U_OrdNum = '" + OrdNum + "'";
					oRecordSet.DoQuery(sQry);

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oMat.RowCount + "건 수정중...!";
				}

				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("검사성적서 납품처수정 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}

				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
		}

		/// <summary>
		/// PS_QM025_CH_CARDCODE
		/// </summary>
		/// <returns></returns>
		private bool PS_QM025_CH_CARDCODE()
		{
			bool ReturnValue = false;
			int i;
			string CardCode;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				CardCode = oMat.Columns.Item("CardCode").Cells.Item(1).Specific.Value.ToString().Trim();
				sQry = "Select CardName From OCRD Where CardCode = '" + CardCode + "'";
				oRecordSet.DoQuery(sQry);

				for (i = 1; i <= oMat.RowCount; i++)
				{
					oMat.Columns.Item("CardCode").Cells.Item(i).Specific.Value = CardCode;
					oMat.Columns.Item("CardName").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
				}

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
		/// PS_QM025_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_QM025_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
				{
					errMessage = "사업장은 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("PackNoF").Specific.Value.ToString().Trim()))
				{
					errMessage = "PACKING 시작번호는 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("PackNoT").Specific.Value.ToString().Trim()))
				{
					errMessage = "PACKING 종료번호는 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}

				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			return ReturnValue;
		}

		
	}
}
