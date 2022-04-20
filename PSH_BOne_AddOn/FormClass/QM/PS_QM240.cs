using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	///  계측기 등록
	/// </summary>
	internal class PS_QM240 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_QM240H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM240L; //등록라인

		private string DocEntry1;
		private int oLast_Mode;

		/// <summary>
		/// Form 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM240.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM240_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM240");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM240_CreateItems();
				PS_QM240_ComboBox_Setting();
				PS_QM240_AddMatrixRow(0, true);
				PS_QM240_LoadCaption();
				PS_QM240_FormReset();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
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
		/// PS_QM240_CreateItems
		/// </summary>
		private void PS_QM240_CreateItems()
		{
			try
			{
				oDS_PS_QM240H = oForm.DataSources.DBDataSources.Item("@PS_QM240H");
				oDS_PS_QM240L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("Cycle", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("Cycle").Specific.DataBind.SetBound(true, "", "Cycle");

				oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM240_ComboBox_Setting
		/// </summary>
		private void PS_QM240_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				dataHelpClass.Set_ComboList(oForm.Items.Item("GCode").Specific, "Select U_Minor = '%', U_CdName = '선택' Union All SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q007' order by U_Minor", "", false, false);
				dataHelpClass.Set_ComboList(oForm.Items.Item("SGCode").Specific, "Select U_Minor = '%', U_CdName = '전체' Union All SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q007' order by U_Minor", "", false, false);

				oForm.Items.Item("GCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("SGCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				dataHelpClass.Set_ComboList(oForm.Items.Item("Sect").Specific, "Select U_Minor = '%', U_CdName = '선택' Union All SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q010' order by U_Minor", "", false, false);
				oForm.Items.Item("Sect").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q007' order by b.U_Minor";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oMat.Columns.Item("GCode").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'Q010' order by b.U_Minor";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oMat.Columns.Item("Sect").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_QM240_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM240_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_QM240L.InsertRecord(oRow);
				}

				oMat.AddRow();
				oDS_PS_QM240L.Offset = oRow;
				oDS_PS_QM240L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 추가 확인 갱신 버튼 이름 변경
		/// </summary>
		private void PS_QM240_LoadCaption()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("Btn_save").Specific.Caption = "추가";
					oForm.Items.Item("Btn_del").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("Btn_save").Specific.Caption = "수정";
					oForm.Items.Item("Btn_del").Enabled = true;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM240_FormReset
		/// </summary>
		private void PS_QM240_FormReset()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Items.Item("GCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("SGCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("Code").Specific.Value = "";
				oForm.Items.Item("CdName").Specific.Value = "";
				oForm.Items.Item("Size").Specific.Value = "";
				oForm.Items.Item("Maker").Specific.Value = "";
				oForm.Items.Item("ProCd").Specific.Value = "";
				oForm.Items.Item("Sect").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("Cycle").Specific.Value = 0;
				oForm.Items.Item("Remark").Specific.Value = "";

				sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM240H]";
				oRecordSet.DoQuery(sQry);
				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					oForm.Items.Item("DocEntry").Specific.Value = 1;
				}
				else
				{
					oForm.Items.Item("DocEntry").Specific.Value = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1;
				}

				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// 조회데이타 가져오기
		/// </summary>
		/// <param name="DocEntry"></param>
		private void PS_QM240_LoadData(string DocEntry)
		{
			int i;
			string SBPLID;
			string SGCode;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				SBPLID = oForm.Items.Item("SBPLId").Specific.Value.ToString().Trim();
				SGCode = oForm.Items.Item("SGCode").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(SGCode))
				{
					SGCode = "%";
				}

				if (DocEntry == "PASS")
				{
					sQry = "EXEC [PS_QM240_01] '" + SBPLID + "','" + SGCode + "','PASS'";
				}
				else
				{
					sQry = "EXEC [PS_QM240_01] '" + SBPLID + "','" + SGCode + "','" + DocEntry + "'";
				}

				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_QM240L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_QM240_AddMatrixRow(0, true);
					PS_QM240_LoadCaption();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				ProgressBar01.Text = "조회시작!";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_QM240L.Size)
					{
						oDS_PS_QM240L.InsertRecord((i));
					}
					oMat.AddRow();
					oDS_PS_QM240L.Offset = i;
					oDS_PS_QM240L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM240L.SetValue("U_ColNum01", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item(4).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item(5).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item(6).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item(7).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item(8).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item(9).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item(10).Value.ToString().Trim());
					oDS_PS_QM240L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item(11).Value.ToString().Trim()).ToString("yyyyMMdd"));
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_QM240_DeleteData
		/// </summary>
		private void PS_QM240_DeleteData()
		{
			string DocEntry;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

					sQry = "Select Count(*) From [@PS_QM240H] where DocEntry = '" + DocEntry + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						errMessage = "삭제대상이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else
					{
						sQry = "Delete From [@PS_QM240H] where DocEntry = '" + DocEntry + "'";
						oRecordSet.DoQuery(sQry);
					}
				}

				PS_QM240_FormReset();
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.Items.Item("Btn_ret").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_QM240_UpdateData
		/// </summary>
		/// <returns></returns>
		private bool PS_QM240_UpdateData()
		{
			bool ReturnValue = false;
			string Size;
			string BPLID;
			string GCode;
			string DocEntry;
			string Code;
			string CdName;
			string Maker;
			string Sect;
			string ProCd;
			string Remark;
			string DocDate;
			int Cycle;
			double d_Cycle;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				GCode = oForm.Items.Item("GCode").Specific.Value.ToString().Trim();
				Code = oForm.Items.Item("Code").Specific.Value.ToString().Trim();
				CdName = oForm.Items.Item("CdName").Specific.Value.ToString().Trim();
				Size = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
				Maker = oForm.Items.Item("Maker").Specific.Value.ToString().Trim();
				ProCd = oForm.Items.Item("ProCd").Specific.Value.ToString().Trim();
				Sect = oForm.Items.Item("Sect").Specific.Value.ToString().Trim();
				// 바로 int Convert시에러
				d_Cycle = Convert.ToDouble(oForm.Items.Item("Cycle").Specific.Value.ToString().Trim());
				Cycle = Convert.ToInt32(d_Cycle);  

				Remark = oForm.Items.Item("Remark").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(DocEntry))
				{
					errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!.";
					throw new Exception();
				}

				sQry = " Update [@PS_QM240H]";
				sQry += " set ";
				sQry += " U_BPLId = '" + BPLID + "',";
				sQry += " U_GCode = '" + GCode + "',";
				sQry += " U_Code = '" + Code + "',";
				sQry += " U_CdName = '" + CdName + "',";
				sQry += " U_Size = '" + Size + "',";
				sQry += " U_Maker = '" + Maker + "',";
				sQry += " U_ProCd  = '" + ProCd + "',";
				sQry += " U_Sect  = '" + Sect + "',";
				sQry += " U_Cycle  = '" + Cycle + "',";
				sQry += " U_Remark = '" + Remark + "'";
				sQry += " Where DocEntry = '" + DocEntry + "'";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.StatusBar.SetText("수정 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
			return ReturnValue;
		}

		/// <summary>
		/// PS_QM240_Add_PurchaseDemand
		/// </summary>
		/// <returns></returns>
		private bool PS_QM240_Add_PurchaseDemand()
		{
			bool ReturnValue = false;
			string Size;
			string Code;
			string BPLID;
			string DocEntry;
			string GCode;
			string CdName;
			string Maker;
			string Sect;
			string ProCd;
			string Remark;
			string DocDate;
			int Cycle;
			double d_Cycle;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				GCode = oForm.Items.Item("GCode").Specific.Value.ToString().Trim();
				Code = oForm.Items.Item("Code").Specific.Value.ToString().Trim();
				CdName = oForm.Items.Item("CdName").Specific.Value.ToString().Trim();
				Size = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
				Maker = oForm.Items.Item("Maker").Specific.Value.ToString().Trim();
				ProCd = oForm.Items.Item("ProCd").Specific.Value.ToString().Trim();
				Sect = oForm.Items.Item("Sect").Specific.Value.ToString().Trim();
				// 바로 int Convert시에러
				d_Cycle = Convert.ToDouble(oForm.Items.Item("Cycle").Specific.Value.ToString().Trim());
				Cycle = Convert.ToInt32(d_Cycle);

				Remark = oForm.Items.Item("Remark").Specific.Value.ToString().Trim();

				sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM240H]";
				oRecordSet.DoQuery(sQry);
				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					DocEntry = "1";
				}
				else
				{
					DocEntry = Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
				}

				sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_Code),3)) + 1)),3) From [@PS_QM240H] Where U_BPLId = '" + BPLID + "' And U_GCode = '" + GCode + "'";
				oRecordSet.DoQuery(sQry);
				if (string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
				{
					Code = BPLID + GCode + "001";
				}
				else
				{
					Code = BPLID + GCode + oRecordSet.Fields.Item(0).Value.ToString().Trim();
				}

				sQry = " INSERT INTO [@PS_QM240H]";
				sQry += " (";
				sQry += " DocEntry,";
				sQry += " DocNum,";
				sQry += " U_BPLId,";
				sQry += " U_DocDate,";
				sQry += " U_GCode,";
				sQry += " U_Code,";
				sQry += " U_CdName,";
				sQry += " U_Size,";
				sQry += " U_Maker,";
				sQry += " U_ProCd,";
				sQry += " U_Sect,";
				sQry += " U_Cycle,";
				sQry += " U_Remark";
				sQry += " ) ";
				sQry += "VALUES(";
				sQry += DocEntry + ",";
				sQry += DocEntry + ",";
				sQry += "'" + BPLID + "',";
				sQry += "'" + DocDate + "',";
				sQry += "'" + GCode + "',";
				sQry += "'" + Code + "',";
				sQry += "'" + CdName + "',";
				sQry += "'" + Size + "',";
				sQry += "'" + Maker + "',";
				sQry += "'" + ProCd + "',";
				sQry += "'" + Sect + "',";
				sQry += "'" + Cycle + "',";
				sQry += "'" + Remark + "'";
				sQry += ")";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				ReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
			return ReturnValue;
		}

		/// <summary>
		/// 입력필수 사항 Check
		/// </summary>
		/// <returns></returns>
		private bool PS_QM240_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;
			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("GCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "분류는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (Convert.ToDouble(oForm.Items.Item("Cycle").Specific.Value.ToString().Trim()) == 0)
				{
					errMessage = "검사주기는 필수사항입니다. 확인하세요.";
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
                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //	Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
					if (pVal.ItemUID == "Btn_save")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_QM240_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_QM240_Add_PurchaseDemand() == false)
							{
								BubbleEvent = false;
								return;
							}

							DocEntry1 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							PS_QM240_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							oLast_Mode = Convert.ToInt32(oForm.Mode);
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM240_UpdateData() == false)
							{
								BubbleEvent = false;
								return;
							}

							DocEntry1 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							PS_QM240_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						}
					}
					else if (pVal.ItemUID == "Btn_ret")
					{
						PS_QM240_LoadData("PASS");
					}
					else if (pVal.ItemUID == "Btn_del")
					{
						PS_QM240_DeleteData();
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Btn_save")
					{
						PS_QM240_LoadCaption();
						PS_QM240_LoadData(DocEntry1);
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			string BPLID;
			string GCode;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "GCode" || pVal.ItemUID == "BPLId")
					{
						BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
						GCode = oForm.Items.Item("GCode").Specific.Value;
						sQry = "Select right('00' + rtrim(Convert(Char(3),Convert(Integer,right(Max(U_Code),3)) + 1)),3) From [@PS_QM240H] Where U_BPLId = '" + BPLID + "' And U_GCode = '" + GCode + "'";
						oRecordSet.DoQuery(sQry);

						if (!string.IsNullOrEmpty(BPLID) && GCode != "%")
						{
							if (string.IsNullOrEmpty(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
							{
								oForm.Items.Item("Code").Specific.String = BPLID + GCode + "001";
							}
							else
							{
								oForm.Items.Item("Code").Specific.String = BPLID + GCode + oRecordSet.Fields.Item(0).Value.ToString().Trim();
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
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
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);

							oForm.Items.Item("DocEntry").Specific.Value = oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("BPLId").Specific.Select(oMat.Columns.Item("BPLId").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("DocDate").Specific.Value = oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("GCode").Specific.Select(oMat.Columns.Item("GCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("SGCode").Specific.Select(oMat.Columns.Item("GCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("Code").Specific.Value = oMat.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("CdName").Specific.Value = oMat.Columns.Item("CdName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("Size").Specific.Value = oMat.Columns.Item("Size").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("Maker").Specific.Value = oMat.Columns.Item("Maker").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("ProCd").Specific.Value = oMat.Columns.Item("ProCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("Sect").Specific.Select(oMat.Columns.Item("Sect").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("Cycle").Specific.Value = oMat.Columns.Item("Cycle").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Items.Item("Remark").Specific.Value = oMat.Columns.Item("Remark").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_QM240_LoadCaption();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM240H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM240L);
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
							PS_QM240_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_QM240_LoadCaption();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
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
