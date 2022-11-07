using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 불합리적출등록
	/// </summary>
	internal class PS_QM070 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_QM070H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM070L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM070.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM070_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM070");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM070_CreateItems();
				PS_QM070_ComboBox_Setting();
				PS_QM070_Add_MatrixRow(0, true);
				PS_QM070_LoadCaption();
				PS_QM070_FormReset();

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
		/// PS_QM070_CreateItems
		/// </summary>
		private void PS_QM070_CreateItems()
		{
			try
			{
				oDS_PS_QM070H = oForm.DataSources.DBDataSources.Item("@PS_QM070H");
				oDS_PS_QM070L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("DocDateF", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.DataSources.UserDataSources.Add("DocDateT", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateF").Specific.DataBind.SetBound(true, "", "DocDateF");
				oForm.Items.Item("DocDateT").Specific.DataBind.SetBound(true, "", "DocDateT");

				oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
				oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 콤보박스 set
		/// </summary>
		private void PS_QM070_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);

				oForm.Items.Item("DocType").Specific.ValidValues.Add("10", "당직");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("20", "일반");
				oForm.Items.Item("DocType").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("ProYN").Specific.ValidValues.Add("N", "미완료");
				oForm.Items.Item("ProYN").Specific.ValidValues.Add("Y", "완료");
				oForm.Items.Item("ProYN").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("SProYN").Specific.ValidValues.Add("N", "미완료");
				oForm.Items.Item("SProYN").Specific.ValidValues.Add("Y", "완료");
				oForm.Items.Item("SProYN").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				sQry = "Select PrcCode, PrcName From [OPRC] Where left(GrpCode,1) in ('1', '2') Order by PrcCode";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("DeptCd").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oMat.Columns.Item("DeptCd").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				sQry = "Select PrcCode, PrcName From [OPRC] Where left(GrpCode,1) in ('1', '2') Order by PrcCode";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("PDeptCd").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oMat.Columns.Item("PDeptCd").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				sQry = "Select PrcCode, PrcName From [OPRC] Where left(GrpCode,1) in ('1', '2') Order by PrcCode";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("SPDeptCd").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//적출코드
				oForm.Items.Item("CR").Specific.ValidValues.Add("", "");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CR").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q002' order by U_Minor", "", false, false);
				oForm.Items.Item("MA").Specific.ValidValues.Add("", "");
				dataHelpClass.Set_ComboList(oForm.Items.Item("MA").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q003' order by U_Minor", "", false, false);
				oForm.Items.Item("MI").Specific.ValidValues.Add("", "");
				dataHelpClass.Set_ComboList(oForm.Items.Item("MI").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q004' order by U_Minor", "", false, false);
				oForm.Items.Item("OB").Specific.ValidValues.Add("", "");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OB").Specific, "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a, [@PS_SY001L] b Where a.Code = b.Code and a.Code = 'Q005' order by U_Minor", "", false, false);

				sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'Q002' Order By U_Minor ";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oMat.Columns.Item("CR").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'Q003' Order By U_Minor ";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oMat.Columns.Item("MA").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'Q004' Order By U_Minor ";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oMat.Columns.Item("MI").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'Q005' Order By U_Minor ";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oMat.Columns.Item("OB").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
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
		/// 메트릭스 Row추가
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM070_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_QM070L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM070L.Offset = oRow;
				oDS_PS_QM070L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		private void PS_QM070_LoadCaption()
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
		/// PS_QM070_FormReset
		/// </summary>
		private void PS_QM070_FormReset()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oDS_PS_QM070H.SetValue("U_CR", 0, "");
				oDS_PS_QM070H.SetValue("U_MA", 0, "");
				oDS_PS_QM070H.SetValue("U_MI", 0, "");
				oDS_PS_QM070H.SetValue("U_OB", 0, "");
				oDS_PS_QM070H.SetValue("U_DeptCd", 0, "");
				oDS_PS_QM070H.SetValue("U_UnitName", 0, "");
				oDS_PS_QM070H.SetValue("U_NGText", 0, "");
				oDS_PS_QM070H.SetValue("U_CntcCode", 0, "");
				oDS_PS_QM070H.SetValue("U_CntcName", 0, "");
				oDS_PS_QM070H.SetValue("U_PDeptCd", 0, "");
				oDS_PS_QM070H.SetValue("U_ProYN", 0, "N");
				oDS_PS_QM070H.SetValue("U_ProDate", 0, "");
				oDS_PS_QM070H.SetValue("U_ProText", 0, "");
				oDS_PS_QM070H.SetValue("U_ProCode", 0, "");
				oDS_PS_QM070H.SetValue("U_ProName", 0, "");
				oDS_PS_QM070H.SetValue("U_Amt", 0, "0");

				sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM070H]";
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// 조회데이타 가져오기
		/// </summary>
		private void PS_QM070_LoadData()
		{
			int i;
			string SProYN;
			string DocDateF;
			string sBPLId;
			string DocDateT;
			string SPDeptCd;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				sBPLId = oForm.Items.Item("SBPLId").Specific.Value.ToString().Trim();
				DocDateF = oForm.Items.Item("DocDateF").Specific.Value.ToString().Trim();
				DocDateT = oForm.Items.Item("DocDateT").Specific.Value.ToString().Trim();
				SProYN = oForm.Items.Item("SProYN").Specific.Value.ToString().Trim();
				SPDeptCd = oForm.Items.Item("SPDeptCd").Specific.Value.ToString().Trim();
				if (string.IsNullOrEmpty(DocDateF))
				{
					DocDateF = "19000101";
				}
				if (string.IsNullOrEmpty(DocDateT))
				{
					DocDateT = "20991231";
				}

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_QM070_01] '" + sBPLId + "','" + DocDateF + "','" + DocDateT + "','" + SProYN + "','" + SPDeptCd + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_QM070L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_QM070_Add_MatrixRow(0, true);
					PS_QM070_LoadCaption();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_QM070L.Size)
					{
						oDS_PS_QM070L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_QM070L.Offset = i;

					oDS_PS_QM070L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM070L.SetValue("U_ColNum01", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item(1).Value.ToString().Trim()); //사업장
					oDS_PS_QM070L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item(2).Value.ToString().Trim()).ToString("yyyyMMdd"));
					oDS_PS_QM070L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item(4).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item(5).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item(6).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item(7).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item(8).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item(9).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item(10).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item(11).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item(12).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColDt02", i, oRecordSet.Fields.Item(13).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg12", i, oRecordSet.Fields.Item(14).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColReg13", i, oRecordSet.Fields.Item(15).Value.ToString().Trim());
					oDS_PS_QM070L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item(16).Value.ToString().Trim());
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
		/// PS_QM070_DeleteData
		/// </summary>
		private void PS_QM070_DeleteData()
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

					sQry = "Select Count(*) From [@PS_QM070H] where DocEntry = '" + DocEntry + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						errMessage = "삭제대상이 없습니다. 확인하세요.";
						throw new Exception();
					}
					else
					{
						sQry = "Delete From [@PS_QM070H] where DocEntry = '" + DocEntry + "'";
						oRecordSet.DoQuery(sQry);
					}
				}

				PS_QM070_FormReset();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_QM070_Updatedata
		/// </summary>
		/// <returns></returns>
		private bool PS_QM070_Updatedata()
		{
			bool ReturnValue = false;
			string CntcCode;
			string UnitName;
			string DeptCd;
			string BPLId;
			string DocEntry;
			string DocType;
			string DeptNm;
			string NGText;
			string CntcName;
			string MI;
			string CR;
			string ProCode;
			string ProDate;
			string PDeptNm;
			string DocDate;
			string PDeptCd;
			string ProYN;
			string ProText;
			string ProName;
			string MA;
			string OB;
			decimal Amt;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocType = oForm.Items.Item("DocType").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				DeptCd = oForm.Items.Item("DeptCd").Specific.Value.ToString().Trim();
				UnitName = oForm.Items.Item("UnitName").Specific.Value.ToString().Trim();
				NGText = oForm.Items.Item("NGText").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
				PDeptCd = oForm.Items.Item("PDeptCd").Specific.Value.ToString().Trim();
				ProYN = oForm.Items.Item("ProYN").Specific.Value.ToString().Trim();
				ProDate = oForm.Items.Item("ProDate").Specific.Value.ToString().Trim();
				ProText = oForm.Items.Item("ProText").Specific.Value.ToString().Trim();
				Amt = Convert.ToDecimal(oForm.Items.Item("Amt").Specific.Value.ToString().Trim());
				ProCode = oForm.Items.Item("ProCode").Specific.Value.ToString().Trim();
				ProName = oForm.Items.Item("ProName").Specific.Value.ToString().Trim();
				CR = oForm.Items.Item("CR").Specific.Value.ToString().Trim();
				MA = oForm.Items.Item("MA").Specific.Value.ToString().Trim();
				MI = oForm.Items.Item("MI").Specific.Value.ToString().Trim();
				OB = oForm.Items.Item("OB").Specific.Value.ToString().Trim();
				DeptNm = "";
				PDeptNm = "";

				if (string.IsNullOrEmpty(DocEntry))
				{
					errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!.";
					throw new Exception();
				}

				sQry = " Update [@PS_QM070H]";
				sQry += " set ";
				sQry += " U_BPLId = '" + BPLId + "',";
				sQry += " U_DocType = '" + DocType + "',";
				sQry += " U_DocDate = '" + DocDate + "',";
				sQry += " U_DeptCd  = '" + DeptCd + "',";
				sQry += " U_DeptNm  = '" + DeptNm + "',";
				sQry += " U_UnitName  = '" + UnitName + "',";
				sQry += " U_NGText  = '" + NGText + "',";
				sQry += " U_CntcCode  = '" + CntcCode + "',";
				sQry += " U_CntcName  = '" + CntcName + "',";
				sQry += " U_PDeptCd  = '" + PDeptCd + "',";
				sQry += " U_PDeptNm  = '" + PDeptNm + "',";
				sQry += " U_ProYN  = '" + ProYN + "',";
				sQry += " U_ProDate  = '" + ProDate + "',";
				sQry += " U_ProText  = '" + ProText + "',";
				sQry += " U_Amt = '" + Amt + "',";
				sQry += " U_CR = '" + CR + "',";
				sQry += " U_MA = '" + MA + "',";
				sQry += " U_MI = '" + MI + "',";
				sQry += " U_OB = '" + OB + "',";
				sQry += " U_ProCode = '" + ProCode + "',";
				sQry += " U_ProName = '" + ProName + "'";
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
		}

		/// <summary>
		/// 데이타 insert
		/// </summary>
		/// <returns></returns>
		private bool PS_QM070_Add_PurchaseDemand()
		{
			bool ReturnValue = false;
			string CntcCode;
			string UnitName;
			string DeptCd;
			string TypeCode;
			string JukCode;
			string BPLId;
			string DocEntry;
			string DocType;
			string JukName;
			string TypeName_Renamed;
			string NGText;
			string CntcName;
			string MI;
			string CR;
			string ProCode;
			string ProDate;
			string DocDate;
			string PDeptCd;
			string ProYN;
			string ProText;
			string ProName;
			string MA;
			string OB;
			decimal Amt;
			string LineNum;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocType = oForm.Items.Item("DocType").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				DeptCd = oForm.Items.Item("DeptCd").Specific.Value.ToString().Trim();
				UnitName = oForm.Items.Item("UnitName").Specific.Value.ToString().Trim();
				NGText = oForm.Items.Item("NGText").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
				PDeptCd = oForm.Items.Item("PDeptCd").Specific.Value.ToString().Trim();
				ProYN = oForm.Items.Item("ProYN").Specific.Value.ToString().Trim();
				ProDate = oForm.Items.Item("ProDate").Specific.Value.ToString().Trim();
				ProText = oForm.Items.Item("ProText").Specific.Value.ToString().Trim();
				Amt = Convert.ToDecimal(oForm.Items.Item("Amt").Specific.Value.ToString().Trim());
				ProCode = oForm.Items.Item("ProCode").Specific.Value.ToString().Trim();
				ProName = oForm.Items.Item("ProName").Specific.Value.ToString().Trim();
				CR = oForm.Items.Item("CR").Specific.Value.ToString().Trim();
				MA = oForm.Items.Item("MA").Specific.Value.ToString().Trim();
				MI = oForm.Items.Item("MI").Specific.Value.ToString().Trim();
				OB = oForm.Items.Item("OB").Specific.Value.ToString().Trim();
				JukCode = "";
				JukName  = "";
				TypeCode = "";
				TypeName_Renamed = "";

				sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_QM070H]";
				oRecordSet.DoQuery(sQry);
				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					DocEntry = "1";
				}
				else
				{
					DocEntry = Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
				}

				sQry = "Select IsNull(Max(U_LineNum), 0) From [@PS_QM070H] Where U_BPLId = '" + BPLId + "' And U_DocDate = '" + DocDate + "'";
				oRecordSet.DoQuery(sQry);
				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					LineNum = "1";
				}
				else
				{
					LineNum = Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
				}

				sQry = "INSERT INTO [@PS_QM070H]";
				sQry += " (";
				sQry += " DocEntry,";
				sQry += " DocNum,";
				sQry += " U_BPLId,";
				sQry += " U_DocType,";
				sQry += " U_DocDate,";
				sQry += " U_LineNum,";
				sQry += " U_JukCode,";
				sQry += " U_JukName,";
				sQry += " U_TypeCode,";
				sQry += " U_TypeName,";
				sQry += " U_DeptCd,";
				sQry += " U_UnitName,";
				sQry += " U_NGText,";
				sQry += " U_CntcCode,";
				sQry += " U_CntcName,";
				sQry += " U_PDeptCd,";
				sQry += " U_ProYN,";
				sQry += " U_ProDate,";
				sQry += " U_ProText,";
				sQry += " U_Amt,";
				sQry += " U_ProCode,";
				sQry += " U_ProName,";
				sQry += " U_CR,";
				sQry += " U_MA,";
				sQry += " U_MI,";
				sQry += " U_OB";
				sQry += " ) ";
				sQry += "VALUES(";
				sQry += DocEntry + ",";
				sQry += DocEntry + ",";
				sQry += "'" + BPLId + "',";
				sQry += "'" + DocType + "',";
				sQry += "'" + DocDate + "',";
				sQry += "'" + LineNum + "',";
				sQry += "'" + JukCode + "',";
				sQry += "'" + JukName + "',";
				sQry += "'" + TypeCode + "',";
				sQry += "'" + TypeName_Renamed + "',";
				sQry += "'" + DeptCd + "',";
				sQry += "'" + UnitName + "',";
				sQry += "'" + NGText + "',";
				sQry += "'" + CntcCode + "',";
				sQry += "'" + CntcName + "',";
				sQry += "'" + PDeptCd + "',";
				sQry += "'" + ProYN + "',";
				sQry += "'" + ProDate + "',";
				sQry += "'" + ProText + "',";
				sQry += "'" + Amt + "',";
				sQry += "'" + ProCode + "',";
				sQry += "'" + ProName + "',";
				sQry += "'" + CR + "',";
				sQry += "'" + MA + "',";
				sQry += "'" + MI + "',";
				sQry += "'" + OB + "'";
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
		}

		/// <summary>
		/// 입력필수 사항 Check
		/// </summary>
		/// <returns></returns>
		private bool PS_QM070_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "입력일자는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("UnitName").Specific.Value.ToString().Trim()))
				{
					errMessage = "적출장소(부위)는 필수사항입니다. 확인하세요.";
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
		/// 코드의 코드명 보여주기
		/// </summary>
		/// <param name="oUID"></param>
		private void PS_QM070_FlushToItemValue(string oUID)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "CntcCode":
						sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
					case "ProCode":
						sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oForm.Items.Item("ProCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("ProName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
				//case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
				//	Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
							if (PS_QM070_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_QM070_Add_PurchaseDemand() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_QM070_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM070_LoadCaption();
							PS_QM070_LoadData();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM070_Updatedata() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_QM070_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_QM070_LoadCaption();
							PS_QM070_LoadData();
							oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
					}
					else if (pVal.ItemUID == "Btn_ret")
					{
						PS_QM070_LoadData();
					}
					else if (pVal.ItemUID == "Btn_del")
					{
						PS_QM070_DeleteData();
						PS_QM070_LoadData();
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "CntcCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "ProCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ProCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
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
		}

		/// <summary>
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string DocEntry;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);
							DocEntry = oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value;
							sQry = "EXEC [PS_QM070_02] '" + DocEntry + "'";
							oRecordSet.DoQuery(sQry);

							if (oRecordSet.RecordCount == 0)
							{
								oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
								PS_QM070_LoadCaption();
								PS_QM070_FormReset();
								errMessage = "결과가 없습니다. 확인하세요.";
								throw new Exception();
							}
							oDS_PS_QM070H.SetValue("DocEntry", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_BPLId", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_DocType", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_DocDate", 0, Convert.ToDateTime(oRecordSet.Fields.Item(3).Value.ToString().Trim()).ToString("yyyyMMdd"));
							oDS_PS_QM070H.SetValue("U_LineNum", 0, oRecordSet.Fields.Item(4).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_CR", 0, oRecordSet.Fields.Item(5).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_MA", 0, oRecordSet.Fields.Item(6).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_MI", 0, oRecordSet.Fields.Item(7).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_OB", 0, oRecordSet.Fields.Item(8).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_DeptCd", 0, oRecordSet.Fields.Item(9).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_UnitName", 0, oRecordSet.Fields.Item(10).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_NGText", 0, oRecordSet.Fields.Item(11).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_CntcCode", 0, oRecordSet.Fields.Item(12).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_CntcName", 0, oRecordSet.Fields.Item(13).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_PDeptCd", 0, oRecordSet.Fields.Item(14).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_ProYN", 0, oRecordSet.Fields.Item(15).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_ProDate", 0, oRecordSet.Fields.Item(16).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_ProText", 0, oRecordSet.Fields.Item(17).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_ProCode", 0, oRecordSet.Fields.Item(18).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_ProName", 0, oRecordSet.Fields.Item(19).Value.ToString().Trim());
							oDS_PS_QM070H.SetValue("U_Amt", 0, oRecordSet.Fields.Item(10).Value.ToString().Trim());

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_QM070_LoadCaption();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
			string DocDate;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.Before_Action == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "CntcCode" || pVal.ItemUID == "ProCode")
						{
							PS_QM070_FlushToItemValue(pVal.ItemUID);
						}
						else if (pVal.ItemUID == "DocDate")
						{
							DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
							sQry = "Select left('" + DocDate + "',6) + '01', Convert(char(8),Dateadd(dd, -1, left(convert(char(8),Dateadd(mm, 1, '" + DocDate + "'),112), 6) + '01'),112)";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("DocDateF").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							oForm.Items.Item("DocDateT").Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
						}
					}
				}
				else if (pVal.Before_Action == false)
				{
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM070H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM070L);
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
							PS_QM070_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_QM070_LoadCaption();
							oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
