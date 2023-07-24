using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 설비정비일지등록
	/// </summary>
	internal class PS_PP190 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.DBDataSource oDS_PS_PP190H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP190L; //등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_PP190M; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oMat01Row01;
		private int oMat02Row02;
		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP190.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP190_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP190");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);

				PS_PP190_CreateItems();
                PS_PP190_SetComboBox();
                PS_PP190_EnableMenus();
                PS_PP190_SetDocument(oFormDocEntry);
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
		/// PS_PP190_CreateItems
		/// </summary>
		private void PS_PP190_CreateItems()
		{
			try
			{
				oDS_PS_PP190H = oForm.DataSources.DBDataSources.Item("@PS_PP190H");
				oDS_PS_PP190L = oForm.DataSources.DBDataSources.Item("@PS_PP190L");
				oDS_PS_PP190M = oForm.DataSources.DBDataSources.Item("@PS_PP190M");

				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("Chk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("Chk").Specific.ValOn = "Y";
				oForm.Items.Item("Chk").Specific.ValOff = "N";
				oForm.Items.Item("Chk").Specific.DataBind.SetBound(true, "", "Chk");
				oForm.DataSources.UserDataSources.Item("Chk").Value = "N";  //미체크로 값을 주고 폼을 로드

				oForm.Items.Item("Year").Specific.Value = DateTime.Now.ToString("yyyy");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP190_SetComboBox
		/// </summary>
		private void PS_PP190_SetComboBox()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				//사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//공정
				sQry = "select U_CpCode, U_CpName from [@PS_PP001L] where U_CpCode in ('CP50108','CP50101') order by 1";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("CpCode").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//금형종류
				sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'P013' order by b.U_Minor";
				oRecordSet.DoQuery(sQry);

				oForm.Items.Item("ToolType").Specific.ValidValues.Add("", "");
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("ToolType").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//구분
				sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'P012' order by b.U_Minor";
				oRecordSet.DoQuery(sQry);

				while (!oRecordSet.EoF)
				{
					oMat01.Columns.Item("Gubun").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oMat01.Columns.Item("Gubun").DisplayDesc = true;

				//Action(Matrix)
				sQry = "  SELECT      U_Minor, ";
				sQry += "             U_CdName ";
				sQry += " FROM        [@PS_SY001L] ";
				sQry += " WHERE       Code = 'A009'";
				sQry += "             AND ISNULL(U_UseYN, 'Y') = 'Y'";
				sQry += " ORDER BY    U_Seq";

				dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Action"), sQry, "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// EnableMenus
		/// </summary>
		private void PS_PP190_EnableMenus()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.SetEnableMenus(oForm, true, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false); //메뉴설정
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP190_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_PP190_SetDocument(string oFormDocEntry)
		{
			int i;
			int sSeq;
			int sCount;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry))
				{
					PS_PP190_EnableFormItem();
					PS_PP190_AddMatrixRow01(0, true);
					PS_PP190_AddMatrixRow02(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_PP190_EnableFormItem();
					oForm.Items.Item("Code").Specific.Value = oFormDocEntry;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("Code").Enabled = false;

					sCount = oMat01.Columns.Item("State").ValidValues.Count;
					sSeq = sCount;
					for (i = 1; i <= sCount; i++)
					{
						oMat01.Columns.Item("State").ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
						sSeq -= 1;
					}

					if (oForm.Items.Item("ToolType").Specific.Value.ToString().Trim() == "3")
					{
						sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'P010' order by b.U_Minor";//금형상태
					}
					else
					{
						sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'P011' order by b.U_Minor";//워크롤상태
					}
					oRecordSet.DoQuery(sQry);

					oMat01.Columns.Item("State").ValidValues.Add("", "");
					while (!oRecordSet.EoF)
					{
						oMat01.Columns.Item("State").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
						oRecordSet.MoveNext();
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP190_AddMatrixRow01
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP190_AddMatrixRow01(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false) //행추가여부
				{
					oDS_PS_PP190L.InsertRecord(oRow);
				}
				oMat01.AddRow();
				oDS_PS_PP190L.Offset = oRow;
				oDS_PS_PP190L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_PP190_AddMatrixRow02
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP190_AddMatrixRow02(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false) //행추가여부
				{
					oDS_PS_PP190M.InsertRecord(oRow);
				}
				oMat02.AddRow();
				oDS_PS_PP190M.Offset = oRow;
				oDS_PS_PP190M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat02.LoadFromDataSource();
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
		/// PS_PP190_CheckDataValid
		/// </summary>
		/// <returns></returns>
		private bool PS_PP190_CheckDataValid()
		{
			bool returnValue = false;
			string errMessage = string.Empty;
			int i;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_PP190H.GetValue("U_BPLId", 0).ToString().Trim()))
				{
					errMessage = "사업장은 필수입력 사항입니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP190H.GetValue("U_CpCode", 0).ToString().Trim()))
				{
					errMessage = "공정은 필수입력 사항입니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP190H.GetValue("U_ToolType", 0).ToString().Trim()))
				{
					errMessage = "금형종류는 필수입력 사항입니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP190H.GetValue("U_LimitCnt", 0).ToString().Trim()))
				{
					errMessage = "한도횟수는 필수입력 사항입니다.";
					throw new Exception();
				}
				else if (string.IsNullOrEmpty(oDS_PS_PP190H.GetValue("U_Year", 0).ToString().Trim()))
				{
					errMessage = "년도는 필수입력 사항입니다.";
					throw new Exception();
				}
				if(oMat01.VisualRowCount == 1)
				{
					errMessage = "라인데이타가 없습니다. 확인하세요.";
					throw new Exception();
				}
				if (oMat02.VisualRowCount == 1)
				{
					errMessage = "호환품목 라인데이타가 없습니다. 확인하세요.";
					throw new Exception();
				}
				for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
				{
					if (oMat01.Columns.Item("Gubun").Cells.Item(i).Specific.Value.ToString().Trim() == "N")
					{
						if (oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value.ToString().Trim() != "1")
						{
							errMessage = "신규입고는 첫번째행에서만 등록가능합니다.";
							oMat01.Columns.Item("Gubun").Cells.Item(i).Specific.Value = "";
							throw new Exception();
						}
					}
					if (string.IsNullOrEmpty(oDS_PS_PP190L.GetValue("U_Gubun", i -1).ToString().Trim()))
					{
						errMessage = "구분는 필수입력사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oDS_PS_PP190L.GetValue("U_HisDate", i - 1).ToString().Trim()))
					{
						errMessage = "이력일자는 필수입력사항입니다. 확인하세요.";
						throw new Exception();
					}
					if(string.IsNullOrEmpty(oDS_PS_PP190L.GetValue("U_State", i - 1).ToString().Trim()))
					{
						errMessage = "상태는 필수입력사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oDS_PS_PP190L.GetValue("U_FinDate", i - 1).ToString().Trim()))
					{
						errMessage = "완료일자는 필수입력사항입니다. 확인하세요.";
						throw new Exception();
					}
					if (Convert.ToDouble(oDS_PS_PP190L.GetValue("U_Thick", i - 1)) <= 0)
					{
						errMessage = "두께는 필수입력사항입니다. 확인하세요.";
						throw new Exception();
					}
				}
				oMat01.FlushToDataSource();
				oMat02.FlushToDataSource();
				oDS_PS_PP190L.RemoveRecord(oDS_PS_PP190L.Size - 1);
				oMat01.LoadFromDataSource();
				oDS_PS_PP190M.RemoveRecord(oDS_PS_PP190M.Size - 1);
				oMat02.LoadFromDataSource();
				returnValue = true;
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
			return returnValue;
		}

		/// <summary>
		/// PS_PP190_EnableFormItem
		/// </summary>
		private void PS_PP190_EnableFormItem()
		{
			int i;
			int sSeq;
			int sCount;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("Seq").Enabled = false;
					oForm.Items.Item("ToolType").Enabled = true;
					oForm.Items.Item("CpCode").Enabled = true;
					oForm.Items.Item("Year").Enabled = true;
					oForm.DataSources.UserDataSources.Item("Chk").Value = "N";
					oMat01.AutoResizeColumns();
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.EnableMenu("1282", true); //추가
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("Seq").Enabled = false;
					oForm.Items.Item("CpCode").Enabled = false;
					oForm.Items.Item("Code").Enabled = true;
					oForm.Items.Item("Seq").Enabled = false;
					oForm.DataSources.UserDataSources.Item("Chk").Value = "N";
					oMat01.AutoResizeColumns();
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.EnableMenu("1282", true); //추가
					oForm.Items.Item("BPLId").Enabled = false;
					oForm.Items.Item("ToolType").Enabled = false;
					oForm.Items.Item("Year").Enabled = false;
					oForm.Items.Item("CpCode").Enabled = false;
					oForm.Items.Item("Code").Enabled = false;
					oForm.Items.Item("Seq").Enabled = false;
					oForm.DataSources.UserDataSources.Item("Chk").Value = "N";
					oMat01.AutoResizeColumns();

					sCount = oMat01.Columns.Item("State").ValidValues.Count;
					sSeq = sCount;
					for (i = 1; i <= sCount; i++)
					{
						oMat01.Columns.Item("State").ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
						sSeq -= 1;
					}
					if (oForm.Items.Item("ToolType").Specific.Value.ToString().Trim() == "3")
					{
						sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'P010' order by b.U_Minor";//금형상태
					}
					else
					{
						sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'P011' order by b.U_Minor"; //워크롤상태
					}
					oRecordSet.DoQuery(sQry);

					oMat01.Columns.Item("State").ValidValues.Add("", "");
					while (!oRecordSet.EoF)
					{
						oMat01.Columns.Item("State").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
						oRecordSet.MoveNext();
					}
				}
				oMat01.AutoResizeColumns();
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
		/// PS_PP190_DeleteAttach
		/// </summary>
		private void PS_PP190_DeleteAttach(int pRow)
		{
			string DeleteFilePath;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat01.FlushToDataSource();
				DeleteFilePath = oDS_PS_PP190L.GetValue("U_AttPath", pRow - 1); //삭제할 첨부파일 경로 저장

				if (string.IsNullOrEmpty(DeleteFilePath))
				{
					errMessage = "첨부파일이 없습니다.";
				}
				else
				{
					if (PSH_Globals.SBO_Application.MessageBox("첨부파일을 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
					{
						System.IO.File.Delete(DeleteFilePath);
						oDS_PS_PP190L.SetValue("U_AttPath", pRow - 1, ""); //첨부파일 경로 삭제
						PSH_Globals.SBO_Application.MessageBox("파일이 삭제되었습니다.");
					}
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
			}
		}

		/// <summary>
		/// PS_PP190_SaveAttach
		/// </summary>
		private void PS_PP190_SaveAttach(int pRow)
		{
			string sFilePath;
			string sFileName;
			string SaveFolders;
			string sourceFile;
			string targetFile;
			string sFileFullPath;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sFileFullPath = PS_PP190_OpenFileSelectDialog();//OpenFileDialog를 쓰레드로 실행

				SaveFolders = "\\\\191.1.1.220\\Attach\\PS_PP190";
				sFileName = System.IO.Path.GetFileName(sFileFullPath); //파일명
				sFilePath = System.IO.Path.GetDirectoryName(sFileFullPath); //파일명을 제외한 전체 경로

				sourceFile = System.IO.Path.Combine(sFilePath, sFileName);
				targetFile = System.IO.Path.Combine(SaveFolders, sFileName);


				if (System.IO.File.Exists(targetFile)) //서버에 기존파일이 존재하는지 체크
				{
					if (PSH_Globals.SBO_Application.MessageBox("동일한 문서번호의 파일이 존재합니다. 교체하시겠습니까?", 2, "Yes", "No") == 1)
					{
						System.IO.File.Delete(targetFile); //삭제
					}
					else
					{
						return;
					}
				}
				oMat01.FlushToDataSource();
				oDS_PS_PP190L.SetValue("U_AttPath", pRow - 1, SaveFolders + "\\" + sFileName); //첨부파일 경로 등록
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();

				System.IO.File.Copy(sourceFile, targetFile, true); //파일 복사 (여기서 오류발생)
				PSH_Globals.SBO_Application.MessageBox("업로드 되었습니다.");

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
			}
		}

		/// <summary>
		/// OpenFileSelectDialog 호출(쓰레드를 이용하여 비동기화)
		/// OLE 호출을 수행하려면 현재 스레드를 STA(단일 스레드 아파트) 모드로 설정해야 합니다.
		/// </summary>
		[STAThread]
		private string PS_PP190_OpenFileSelectDialog()
		{
			string returnFileName = string.Empty;

			var thread = new System.Threading.Thread(() =>
			{
				System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
				openFileDialog.InitialDirectory = "C:\\";
				openFileDialog.Filter = "All files (*.*)|*.*";
				openFileDialog.FilterIndex = 1; //FilterIndex는 1부터 시작
				openFileDialog.RestoreDirectory = true;

				if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
				{
					returnFileName = openFileDialog.FileName;
				}
			});

			thread.SetApartmentState(System.Threading.ApartmentState.STA);
			thread.Start();
			thread.Join();
			return returnFileName;
		}

		/// <summary>
		/// PS_PP190_OpenAttach
		/// </summary>
		private void PS_PP190_OpenAttach(int pRow)
		{
			string AttachPath;
			string errMessage = string.Empty;

			try
			{
				//oMat01.FlushToDataSource();
				AttachPath = oDS_PS_PP190L.GetValue("U_AttPath", pRow - 1).ToString().Trim();
				if (string.IsNullOrEmpty(AttachPath))
				{
					PSH_Globals.SBO_Application.MessageBox("첨부파일이 없습니다.");
				}
				else
				{
					System.Diagnostics.ProcessStartInfo process = new System.Diagnostics.ProcessStartInfo(AttachPath);
					process.UseShellExecute = true;
					process.Verb = "open";

					System.Diagnostics.Process.Start(process);
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
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
				if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02")
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
				if (pVal.ItemUID == "Mat01")
				{
					if (pVal.Row > 0)
					{
						oMat01Row01 = pVal.Row;
					}
				}
				else if (pVal.ItemUID == "Mat02")
				{
					if (pVal.Row > 0)
					{
						oMat02Row02 = pVal.Row;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
			string cLen;
			string sQry;
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_PP190_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (oForm.DataSources.UserDataSources.Item("Chk").Value.ToString().Trim() == "N")
							{
								sQry = "Select ISNULL(MAX(U_Seq),0) + 1";
								sQry += "From [@PS_PP190H] ";
								sQry += "Where U_ToolType = '" + oForm.Items.Item("ToolType").Specific.Value.ToString().Trim() + "' ";
								sQry += "And U_Year = '" + oForm.Items.Item("Year").Specific.Value.ToString().Trim() + "'";
								sQry += "And U_BPLId = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 100)
								{
									PSH_Globals.SBO_Application.SetStatusBarMessage("순번이 99를 초과할 수 없습니다. 관리자에게 문의하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
								}
								oForm.Items.Item("Seq").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().PadLeft(2, '0');
								oForm.Items.Item("Code").Specific.Value = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()
																		  + oForm.Items.Item("ToolType").Specific.Value.ToString().Trim()
																		  + codeHelpClass.Right(oForm.Items.Item("Year").Specific.Value.ToString().Trim(), 2)
																		  + oRecordSet.Fields.Item(0).Value.ToString().PadLeft(2, '0');
							}
							else
							{
								cLen = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()
									   + oForm.Items.Item("ToolType").Specific.Value.ToString().Trim()
									  + codeHelpClass.Right(oForm.Items.Item("Year").Specific.Value.ToString().Trim(), 2)
									  + oForm.Items.Item("Seq").Specific.Value.ToString().Trim();

								if (cLen.Length != 6)
								{
									PSH_Globals.SBO_Application.SetStatusBarMessage("코드가 6자리여야 합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
									BubbleEvent = false;
									return;
								}
								oForm.Items.Item("Code").Specific.Value = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()
																		 + oForm.Items.Item("ToolType").Specific.Value.ToString().Trim()
																		 + codeHelpClass.Right(oForm.Items.Item("Year").Specific.Value.ToString().Trim(), 2)
																		 + oForm.Items.Item("Seq").Specific.Value.ToString().Trim();
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_PP190_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
					//취소버튼 누를시 저장할 자료가 있으면 메시지 표시
					if (pVal.ItemUID == "2")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (oMat01.VisualRowCount > 1)
							{
								if (PSH_Globals.SBO_Application.MessageBox("저장하지 않는 자료가 있습니다. 취소하시겠습니까?", 1, "예", "아니오") != 1)
								{
									BubbleEvent = false;
									return;
								}
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_PP190_EnableFormItem();
								PS_PP190_AddMatrixRow01(0, true);
								PS_PP190_AddMatrixRow02(0, true);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								if (oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
								{
									PS_PP190_EnableFormItem();
									oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
									oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}
								PS_PP190_EnableFormItem();
							}
						}
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
							oMat01Row01 = pVal.Row;
						}
					}
					if (pVal.ItemUID == "Mat02")
					{
						if (pVal.Row > 0)
						{
							oMat02.SelectRow(pVal.Row, true, false);
							oMat02Row02 = pVal.Row;
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
			int i;
			int sSeq;
			int sCount;
			string sQry;
			string errMessage = string.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "ToolType")
					{
						sCount = oMat01.Columns.Item("State").ValidValues.Count;
						sSeq = sCount;
						for (i = 1; i <= sCount; i++)
						{
							oMat01.Columns.Item("State").ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
							sSeq -= 1;
						}
					}
					else if (pVal.ItemUID == "CpCode")
					{
						sCount = oForm.Items.Item("ToolType").Specific.ValidValues.Count;
						sSeq = sCount;
						for (i = 1; i <= sCount; i++)
						{
							oForm.Items.Item("ToolType").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
							sSeq -= 1;
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "ToolType")
					{
						if (oForm.Items.Item("ToolType").Specific.Value.ToString().Trim() == "3")
						{
							sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'P010' order by b.U_Minor";//금형상태
						}
						else
						{
							sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'P011' order by b.U_Minor"; //워크롤상태
						}
						oRecordSet.DoQuery(sQry);

						oMat01.Columns.Item("State").ValidValues.Add("", "");
						while (!oRecordSet.EoF)
						{
							oMat01.Columns.Item("State").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oRecordSet.MoveNext();
						}
					}
					else if (pVal.ItemUID == "CpCode")
					{
						sQry = "SELECT b.U_Minor, b.U_CdName FROM [@PS_SY001H] a Inner Join [@PS_SY001L] b On a.Code = b.Code And a.Code = 'P013' and b.U_RelCd = '";
						sQry += oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() + "' order by b.U_Minor";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("ToolType").Specific.ValidValues.Add("", "");
						while (!oRecordSet.EoF)
						{
							oForm.Items.Item("ToolType").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oRecordSet.MoveNext();
						}
						oForm.Items.Item("ToolType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_ByValue);
					}

					if (pVal.ItemUID == "Mat01")
					{
						oMat01.FlushToDataSource();
						if (pVal.ColUID == "Action")
						{
							if (oMat01.Columns.Item("Action").Cells.Item(pVal.Row).Specific.Value == "S")
							{
								PS_PP190_SaveAttach(pVal.Row);
							}
							else if (oMat01.Columns.Item("Action").Cells.Item(pVal.Row).Specific.Value == "O")
                            {
                                PS_PP190_OpenAttach(pVal.Row);
                            }
                            else if (oMat01.Columns.Item("Action").Cells.Item(pVal.Row).Specific.Value == "D")
                            {
                                PS_PP190_DeleteAttach(pVal.Row);
                            }
						}
						else if (pVal.ColUID == "Gubun")
						{
							if (pVal.ItemChanged == true)
							{
								if (oDS_PS_PP190L.GetValue("U_Gubun", pVal.Row - 1).ToString().Trim() != "N" && oMat01.VisualRowCount == pVal.Row)
								{
									oForm.Items.Item("UseCnt").Specific.Value = "0";
								}
								oMat01.FlushToDataSource();
								if(string.IsNullOrEmpty(oDS_PS_PP190L.GetValue("U_HisDate", pVal.Row - 1).ToString().Trim()))
								{
									oDS_PS_PP190L.SetValue("U_HisDate", pVal.Row - 1, DateTime.Now.ToString("yyyyMMdd"));
									oDS_PS_PP190L.SetValue("U_FinDate", pVal.Row - 1, DateTime.Now.ToString("yyyyMMdd"));
								}
								oDS_PS_PP190L.SetValue("U_Thick", pVal.Row - 1, "0");
								oDS_PS_PP190L.SetValue("U_CntcCode", pVal.Row - 1, dataHelpClass.User_MSTCOD());
								oDS_PS_PP190L.SetValue("U_CntcName", pVal.Row - 1, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + dataHelpClass.User_MSTCOD() + "'", ""));
								if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP190L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_PP190_AddMatrixRow01(pVal.Row, false);
								}
								oMat01.LoadFromDataSource();
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				if(errMessage != string.Empty)
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
				oForm.Freeze(false);
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
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
							oMat01.FlushToDataSource();
							if (pVal.ColUID == "CntcCode")
							{
								oDS_PS_PP190L.SetValue("U_CntcName", pVal.Row - 1, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));
							}
						}
						else if (pVal.ItemUID == "Mat02")
						{
							oMat02.FlushToDataSource();
							if (pVal.ColUID == "ItemCode")
							{
								oDS_PS_PP190M.SetValue("U_ItemName", pVal.Row - 1, dataHelpClass.Get_ReData("ItemName", "ItemCode", "OITM", "'" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));
								if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP190M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_PP190_AddMatrixRow02(pVal.Row, false);
								}
							}
							else
							{
								oDS_PS_PP190M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							}
						}
						oMat01.LoadFromDataSource();
						oMat02.LoadFromDataSource();
						oMat01.AutoResizeColumns();
						oMat02.AutoResizeColumns();
						oForm.Update();

						if (pVal.ItemUID == "Mat01")
						{
							oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else if (pVal.ItemUID == "Mat02")
						{
							oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else
						{
							oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_MATRIX_LOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP190_EnableFormItem();
					if (pVal.ItemUID == "Mat01")
					{
						PS_PP190_AddMatrixRow01(oMat01.VisualRowCount, false);
					}
					else if (pVal.ItemUID == "Mat02")
					{
						PS_PP190_AddMatrixRow02(oMat02.VisualRowCount, false);
					}
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
				if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02")
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
				if (pVal.ItemUID == "Mat01")
				{
					if (pVal.Row > 0)
					{
						oMat01Row01 = pVal.Row;
					}
				}
				else if (pVal.ItemUID == "Mat02")
				{
					if (pVal.Row > 0)
					{
						oMat02Row02 = pVal.Row;
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP190H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP190L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP190M);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
						//행삭제전 행삭제가능여부검사
					}
					else if (pVal.BeforeAction == false)
					{
						if (oLastItemUID01 == "Mat01")
						{
							for (i = 1; i <= oMat01.VisualRowCount; i++)
							{
								oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
							}
							oMat01.FlushToDataSource();
							oDS_PS_PP190L.RemoveRecord(oDS_PS_PP190L.Size - 1);
							oMat01.LoadFromDataSource();
							if (oMat01.RowCount == 0)
							{
								PS_PP190_AddMatrixRow01(0, false);
							}
							else
							{
								if (!string.IsNullOrEmpty(oDS_PS_PP190L.GetValue("U_HisDate", oMat01.RowCount - 1).ToString().Trim()))
								{
									PS_PP190_AddMatrixRow01(oMat01.RowCount, false);
								}
							}
						}
						else if (oLastItemUID01 == "Mat02")
						{
							for (i = 1; i <= oMat02.VisualRowCount; i++)
							{
								oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
							}
							oMat02.FlushToDataSource();
							oDS_PS_PP190M.RemoveRecord(oDS_PS_PP190M.Size - 1);
							oMat02.LoadFromDataSource();
							if (oMat02.RowCount == 0)
							{
								PS_PP190_AddMatrixRow02(0, false);
							}
							else
							{
								if (!string.IsNullOrEmpty(oDS_PS_PP190M.GetValue("U_ItemCode", oMat02.RowCount - 1).ToString().Trim()))
								{
									PS_PP190_AddMatrixRow02(oMat02.RowCount, false);
								}
							}
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
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
							{
								if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
								{
									BubbleEvent = false;
									return;
								}
							}
							break;
						case "1286": //닫기
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_PP190_EnableFormItem();
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							PS_PP190_EnableFormItem();
							break;
						case "1282": //추가
							PS_PP190_EnableFormItem();
							PS_PP190_AddMatrixRow01(0, true);
							PS_PP190_AddMatrixRow02(0, true);
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_PP190_EnableFormItem();
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
