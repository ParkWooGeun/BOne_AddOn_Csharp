using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 월별계획등록
	/// </summary>
	internal class PS_GA032 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_GA032A; //등록용 Matrix

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_GA032.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_GA032_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_GA032");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_GA032_CreateItems();
				PS_GA032_ComboBox_Setting();
				PS_GA032_LoadCaption();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);

				PS_GA032_FormReset(); //폼초기화
				oForm.Items.Item("StdYear").Click();
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
		/// PS_GA032_CreateItems
		/// </summary>
		private void PS_GA032_CreateItems()
		{
			try
			{
				oDS_PS_GA032A = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//기준년도
				oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
				oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");
				oForm.Items.Item("StdYear").Specific.Value = DateTime.Now.ToString("yyyy");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 콤보박스 set
		/// </summary>
		private void PS_GA032_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//계획구분
				oMat.Columns.Item("PlanCls").ValidValues.Add("01", "최초계획");
				oMat.Columns.Item("PlanCls").ValidValues.Add("02", "수정계획");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
		/// </summary>
		private void PS_GA032_LoadCaption()
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// 화면 초기화
		/// </summary>
		private void PS_GA032_FormReset()
		{
			string User_BPLId;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				User_BPLId = dataHelpClass.User_BPLID();
				oForm.DataSources.UserDataSources.Item("StdYear").Value = DateTime.Now.ToString("yyyy");
				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();
				PS_GA032_Add_MatrixRow(0, true);
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
		/// 메트릭스 Row추가
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_GA032_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_GA032A.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_GA032A.Offset = oRow;
				oDS_PS_GA032A.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oDS_PS_GA032A.SetValue("U_ColReg01", oRow, "Y");
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_GA032_CheckAll
		/// </summary>
		private void PS_GA032_CheckAll()
		{
			int loopCount;
			string CheckType;

			try
			{
				oForm.Freeze(true);
				CheckType = "Y";

				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_GA032A.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break; // TODO: might not be correct. Was : Exit For
					}
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					oDS_PS_GA032A.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_GA032A.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_GA032A.SetValue("U_ColReg01", loopCount, "N");
					}
				}

				oMat.LoadFromDataSource();
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
		/// 등록된 데이터 조회
		/// </summary>
		private void PS_GA032_MTX01()
		{
			int i;
			string BPLID;
			string StdYear;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "      EXEC [PS_GA032_01] '";
				sQry += BPLID + "','";
				sQry += StdYear + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_GA032A.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_GA032_Add_MatrixRow(0, true);
					PS_GA032_LoadCaption();
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_GA032A.Size)
					{
						oDS_PS_GA032A.InsertRecord(i);
					}
					oMat.AddRow();
					oDS_PS_GA032A.Offset = i;
					oDS_PS_GA032A.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_GA032A.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Select").Value.ToString().Trim());                  //선택
					oDS_PS_GA032A.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("BudCls").Value.ToString().Trim());                  //구분
					oDS_PS_GA032A.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("BudClsNm").Value.ToString().Trim());                    //구분명
					oDS_PS_GA032A.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("PlanCls").Value.ToString().Trim());                 //계획구분
					oDS_PS_GA032A.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Month01").Value.ToString().Trim());                 //1월
					oDS_PS_GA032A.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("Month02").Value.ToString().Trim());                 //2월
					oDS_PS_GA032A.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("Month03").Value.ToString().Trim());                 //3월
					oDS_PS_GA032A.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("Month04").Value.ToString().Trim());                 //4월
					oDS_PS_GA032A.SetValue("U_ColSum05", i, oRecordSet.Fields.Item("Month05").Value.ToString().Trim());                 //5월
					oDS_PS_GA032A.SetValue("U_ColSum06", i, oRecordSet.Fields.Item("Month06").Value.ToString().Trim());                 //6월
					oDS_PS_GA032A.SetValue("U_ColSum07", i, oRecordSet.Fields.Item("Month07").Value.ToString().Trim());                 //7월
					oDS_PS_GA032A.SetValue("U_ColSum08", i, oRecordSet.Fields.Item("Month08").Value.ToString().Trim());                 //8월
					oDS_PS_GA032A.SetValue("U_ColSum09", i, oRecordSet.Fields.Item("Month09").Value.ToString().Trim());                 //9월
					oDS_PS_GA032A.SetValue("U_ColSum10", i, oRecordSet.Fields.Item("Month10").Value.ToString().Trim());                 //10월
					oDS_PS_GA032A.SetValue("U_ColSum11", i, oRecordSet.Fields.Item("Month11").Value.ToString().Trim());                 //11월
					oDS_PS_GA032A.SetValue("U_ColSum12", i, oRecordSet.Fields.Item("Month12").Value.ToString().Trim());                 //12월
					oDS_PS_GA032A.SetValue("U_ColSum13", i, oRecordSet.Fields.Item("Total").Value.ToString().Trim());                   //계
					oRecordSet.MoveNext();

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				PS_GA032_Add_MatrixRow(oMat.VisualRowCount, false);
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
		/// 기본정보 삭제
		/// </summary>
		public void PS_GA032_DeleteData()
		{
			int loopCount;
			string BPLID;
			string StdYear;
			string BudCls;
			string PlanCls;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.DataSources.UserDataSources.Item("BPLID").Value.ToString().Trim();
				StdYear = oForm.DataSources.UserDataSources.Item("StdYear").Value.ToString().Trim();

				oMat.FlushToDataSource();

				//마지막 빈행 제외를 위해 2를 뺌
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 2; loopCount++)
				{
					if (oDS_PS_GA032A.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						BudCls = oDS_PS_GA032A.GetValue("U_ColReg02", loopCount).ToString().Trim();
						PlanCls = oDS_PS_GA032A.GetValue("U_ColReg04", loopCount).ToString().Trim();

						sQry = "      EXEC [PS_GA032_03] '";
						sQry += BPLID + "','";
						sQry += StdYear + "','";
						sQry += BudCls + "','";
						sQry += PlanCls + "'";
						oRecordSet.DoQuery(sQry);
					}
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
		/// </summary>
		/// <returns></returns>
		private bool PS_GA032_AddData()
		{
			bool ReturnValue = false;
			int i;
			string BPLID;    //사업장
			string StdYear;  //기준년도
			string BudCls;   //구분
			string PlanCls;  //계획구분
			decimal Month01; //1월
			decimal Month02; //2월
			decimal Month03; //3월
			decimal Month04; //4월
			decimal Month05; //5월
			decimal Month06; //6월
			decimal Month07; //7월
			decimal Month08; //8월
			decimal Month09; //9월
			decimal Month10; //10월
			decimal Month11; //11월
			decimal Month12; //12월
			decimal Total;   //계
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLID = oForm.DataSources.UserDataSources.Item("BPLID").Value;
				StdYear = oForm.DataSources.UserDataSources.Item("StdYear").Value;

				ProgressBar01.Text = "저장중.....";

				oMat.FlushToDataSource();
				//마지막 빈행 제외를 위해 2를 뺌
				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					if (oDS_PS_GA032A.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
					{
						BudCls = oDS_PS_GA032A.GetValue("U_ColReg02", i).ToString().Trim();  //구분
						PlanCls = oDS_PS_GA032A.GetValue("U_ColReg04", i).ToString().Trim(); //계획구분
						Month01 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum01", i).ToString().Trim());
						Month02 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum02", i).ToString().Trim());
						Month03 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum03", i).ToString().Trim());
						Month04 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum04", i).ToString().Trim());
						Month05 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum05", i).ToString().Trim());
						Month06 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum06", i).ToString().Trim());
						Month07 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum07", i).ToString().Trim());
						Month08 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum08", i).ToString().Trim());
						Month09 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum09", i).ToString().Trim());
						Month10 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum10", i).ToString().Trim());
						Month11 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum11", i).ToString().Trim());
						Month12 = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum12", i).ToString().Trim());
						Total = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum13", i).ToString().Trim());

						sQry = " EXEC [PS_GA032_02] '";
						sQry += BPLID + "','";
						sQry += StdYear + "','";
						sQry += BudCls + "','";
						sQry += PlanCls + "','";
						sQry += Month01 + "','";
						sQry += Month02 + "','";
						sQry += Month03 + "','";
						sQry += Month04 + "','";
						sQry += Month05 + "','";
						sQry += Month06 + "','";
						sQry += Month07 + "','";
						sQry += Month08 + "','";
						sQry += Month09 + "','";
						sQry += Month10 + "','";
						sQry += Month11 + "','";
						sQry += Month12 + "','";
						sQry += Total + "'";
						oRecordSet.DoQuery(sQry);

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToString(oMat.VisualRowCount - 1) + "건 저장중...";
					}
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
		/// PS_GA032_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_GA032_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string FldValue;
			string Descr;
			decimal Total;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "Mat01":
						oMat.FlushToDataSource();

						if (oCol == "BudCls")
						{
							//예산구분 조회
							sQry = " SELECT      T0.FldValue,";
							sQry += "             T0.Descr";
							sQry += " FROM        UFD1 AS T0";
							sQry += " WHERE       T0.TableID = 'JDT1'";
							sQry += "             AND T0.FieldID = '28'";
							sQry += "             AND T0.FldValue = '" + oMat.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							FldValue = oRecordSet.Fields.Item("FldValue").Value.ToString().Trim();
							Descr = oRecordSet.Fields.Item("Descr").Value.ToString().Trim();

							oDS_PS_GA032A.SetValue("U_ColReg02", oRow - 1, FldValue); //예산구분코드
							oDS_PS_GA032A.SetValue("U_ColReg03", oRow - 1, Descr);	  //예산구분명

							if (oMat.RowCount == oRow && !string.IsNullOrEmpty(oDS_PS_GA032A.GetValue("U_ColReg02", oRow - 1).ToString().Trim())) 
							{
								PS_GA032_Add_MatrixRow(oRow, false);
							}
						}
						else if (oCol != "BudCls")
						{
							Total = Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum01", oRow - 1).ToString().Trim())
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum02", oRow - 1).ToString().Trim()) 
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum03", oRow - 1).ToString().Trim()) 
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum04", oRow - 1).ToString().Trim()) 
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum05", oRow - 1).ToString().Trim()) 
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum06", oRow - 1).ToString().Trim()) 
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum07", oRow - 1).ToString().Trim()) 
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum08", oRow - 1).ToString().Trim())
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum09", oRow - 1).ToString().Trim()) 
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum10", oRow - 1).ToString().Trim()) 
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum11", oRow - 1).ToString().Trim())
								  + Convert.ToDecimal(oDS_PS_GA032A.GetValue("U_ColSum12", oRow - 1).ToString().Trim());

							oDS_PS_GA032A.SetValue("U_ColSum13", oRow - 1, Convert.ToString(Total));
						}

						oMat.LoadFromDataSource();
						oMat.Columns.Item(oCol).Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oMat.AutoResizeColumns();
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
				//	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
					if (pVal.ItemUID == "BtnAdd")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_GA032_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_GA032_LoadCaption();
							PS_GA032_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							PS_GA032_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_GA032_LoadCaption();
							PS_GA032_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSearch")
					{
						PS_GA032_MTX01();
					}
					else if (pVal.ItemUID == "BtnDelete")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
						{
							PS_GA032_DeleteData();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_GA032_LoadCaption();
							PS_GA032_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSelect")
					{
						PS_GA032_CheckAll();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "BudCls");
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
							PS_GA032_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							oMat.AutoResizeColumns();
						}
						else
						{
							PS_GA032_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_GA032A);
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
						case "1281": //찾기
							break;
						case "1282": //추가
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_GA032_FormReset();
							break;
						case "1283": //삭제
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
							break;
						case "7169": //엑셀 내보내기
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1287": // 복제
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
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
