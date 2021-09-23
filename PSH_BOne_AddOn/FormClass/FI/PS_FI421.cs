using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 입금표등록
	/// </summary>
	internal class PS_FI421 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_FI421H;// 등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_FI421L;// 등록라인
		private string oLastItemUID01;  // 클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;   // 마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;      // 마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Mode;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();
			
			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI421.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI421_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI421");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_FI421_CreateItems();
				PS_FI421_ComboBox_Setting();
				PS_FI421_EnableMenus();
				PS_FI421_FormResize();
				PS_FI421_Add_MatrixRow(0, true);
				PS_FI421_LoadCaption();
				PS_FI421_FormReset();
				
				oDS_PS_FI421H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd")); // 발행일자 설정
				oForm.Items.Item("SDocDateFr").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01"; // 발행일자FR
				oForm.Items.Item("SDocDateTo").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); // 발행일자TO
				oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular); // 사번 포커스
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc01); //메모리 해제
			}
		}

		/// <summary>
		/// PS_FI421_LoadCaption
		/// </summary>
		private void PS_FI421_LoadCaption()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
					oForm.Items.Item("BtnDelete").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
					oForm.Items.Item("BtnDelete").Enabled = true;
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
		/// PS_FI421_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_FI421_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)   // 행추가여부
				{
					oDS_PS_FI421L.InsertRecord(oRow);
				}
				oMat01.AddRow();
				oDS_PS_FI421L.Offset = oRow;
				oDS_PS_FI421L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat01.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// PS_FI421_MTX01
		/// </summary>
		private void PS_FI421_MTX01()
		{
			int i = 0;
			int ErrNum = 0;
			string sQry = string.Empty;
			string sDocEntry = string.Empty;			//관리번호
			string SSerialNo = string.Empty;			//일련번호
			string SBPLID = string.Empty;			    //사업장
			string SRspCode = string.Empty;			    //담당
			string SCntcCode = string.Empty;			//사번
			string SCardCode = string.Empty;			//거래처
			string SDocDateFr = string.Empty;			//발행일자
			string SDocDateTo = string.Empty;			//발행일자TO
			decimal SAmount = 0;	                    //공급가액
			decimal SVatTax = 0;	                    //부가가치세
			string SContents = string.Empty;			//내용

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				sDocEntry = oForm.Items.Item("SDocEntry").Specific.Value.ToString().Trim();            //관리번호
				SSerialNo = oForm.Items.Item("SSerialNo").Specific.Value.ToString().Trim();               //일련번호
				SBPLID = oForm.Items.Item("SBPLId").Specific.Value.ToString().Trim();                  //사업장
				SRspCode = oForm.Items.Item("SRspCode").Specific.Value.ToString().Trim();                //담당
				SCntcCode = oForm.Items.Item("SCntcCode").Specific.Value.ToString().Trim();               //사번
				SCardCode = oForm.Items.Item("SCardCode").Specific.Value.ToString().Trim();               //거래처
				SDocDateFr = oForm.Items.Item("SDocDateFr").Specific.Value.ToString().Trim();              //발행일자FR
				SDocDateTo = oForm.Items.Item("SDocDateTo").Specific.Value.ToString().Trim();              //발행일자TO
				SAmount = Convert.ToDecimal(oForm.Items.Item("SAmount").Specific.Value.ToString().Trim()); //공급가액
				SVatTax = Convert.ToDecimal(oForm.Items.Item("SVatTax").Specific.Value.ToString().Trim()); //부가가치세
				SContents = oForm.Items.Item("SContents").Specific.Value.ToString().Trim();               //내용

				sQry = "EXEC [PS_FI421_01] '";
				sQry += sDocEntry + "','";
				sQry += SSerialNo + "','";
				sQry += SBPLID + "','";
				sQry += SRspCode + "','";
				sQry += SCntcCode + "','";
				sQry += SCardCode + "','";
				sQry += SDocDateFr + "','";
				sQry += SDocDateTo + "','";
				sQry += SAmount + "','";
				sQry += SVatTax + "','";
				sQry += SContents + "'";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_FI421L.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					ErrNum = 1;
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_FI421_Add_MatrixRow(0, true);
					PS_FI421_LoadCaption();
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_FI421L.Size)
					{
						oDS_PS_FI421L.InsertRecord(i);
					}
					oMat01.AddRow();
					oDS_PS_FI421L.Offset = i;
					oDS_PS_FI421L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_FI421L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());				//관리번호
					oDS_PS_FI421L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("SerialNo").Value.ToString().Trim());				//일련번호
					oDS_PS_FI421L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("BPLId").Value.ToString().Trim());					//사업장
					oDS_PS_FI421L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("RspCode").Value.ToString().Trim());					//담당
					oDS_PS_FI421L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("CntcCode").Value.ToString().Trim());				//사번
					oDS_PS_FI421L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("CntcName").Value.ToString().Trim());				//성명
					oDS_PS_FI421L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());				//거래처
					oDS_PS_FI421L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());				//거래처명
					oDS_PS_FI421L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd"));     //발행일자
					oDS_PS_FI421L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Amount").Value.ToString().Trim());					//공급가액
					oDS_PS_FI421L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("VatTax").Value.ToString().Trim());					//부가가치세
					oDS_PS_FI421L.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("Contents").Value.ToString().Trim());				//내용

					oRecordSet.MoveNext();
					ProgBar01.Value = ProgBar01.Value + 1;
					ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.MessageBox("조회 결과가 없습니다. 확인하세요.");
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				ProgBar01.Stop();
				oForm.Freeze(false);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
			}
		}

		/// <summary>
		/// PS_FI421_DeleteData
		/// </summary>
		private void PS_FI421_DeleteData()
		{
			int ErrNum = 0;
			string sQry = string.Empty;
			string DocEntry = string.Empty;
			
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

					sQry = "SELECT COUNT(*) FROM [@PS_FI421H] WHERE DocEntry = '" + DocEntry + "'";
					oRecordSet.DoQuery(sQry);

					if (oRecordSet.RecordCount == 0)
					{
						ErrNum = 1;
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						throw new Exception();
					}
					else
					{
						sQry = "DELETE FROM [@PS_FI421H] WHERE DocEntry = '" + DocEntry + "'";
						oRecordSet.DoQuery(sQry);
					}
				}
				PSH_Globals.SBO_Application.MessageBox("삭제 완료!");
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("삭제대상이 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
		/// PS_FI421_UpdateData
		/// </summary>
		/// <returns></returns>
		private bool PS_FI421_UpdateData()
		{
			bool functionReturnValue = false;

			short DocEntry = 0;
			string SerialNo = string.Empty;			//일련번호
			string BPLID = string.Empty;			//사업장
			string RspCode = string.Empty;			//담당
			string CntcCode = string.Empty;			//사번
			string CntcName = string.Empty;			//성명
			string CardCode = string.Empty;			//거래처
			string CardName = string.Empty;			//거래처명
			string DocDate = string.Empty;			//발행일자
			decimal Amount = 0;		            	//공급가액
			decimal VatTax = 0;		            	//부가가치세
			string Contents = string.Empty;			//내용
			string UserSign = string.Empty;         //UserSign
			string sQry = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = Convert.ToInt16(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());        
				SerialNo = oForm.Items.Item("SerialNo").Specific.Value.ToString().Trim();            
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();                  
				RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();              
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();            
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();            
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();            
				CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();            
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();              
				Amount = Convert.ToDecimal(oForm.Items.Item("Amount").Specific.Value.ToString().Trim());
				VatTax = Convert.ToDecimal(oForm.Items.Item("VatTax").Specific.Value.ToString().Trim());
				Contents = oForm.Items.Item("Contents").Specific.Value.ToString().Trim(); 
				UserSign = PSH_Globals.oCompany.UserSignature.ToString();
				
				if (string.IsNullOrEmpty(Convert.ToString(DocEntry)))
				{
					PSH_Globals.SBO_Application.MessageBox("수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!");
					throw new Exception();
				}

				sQry = "EXEC [PS_FI421_03] '";
				sQry += DocEntry + "','";
				sQry += SerialNo + "','";
				sQry += BPLID + "','";
				sQry += RspCode + "','";
				sQry += CntcCode + "','";
				sQry += CntcName + "','";
				sQry += CardCode + "','";
				sQry += CardName + "','";
				sQry += DocDate + "','";
				sQry += Amount + "','";
				sQry += VatTax + "','";
				sQry += Contents + "'";

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
		/// PS_FI421_AddData
		/// </summary>
		/// <returns></returns>
		private bool PS_FI421_AddData()
		{
			bool functionReturnValue = false;

			string sQry = string.Empty;

			double DocEntry = 0;
			string SerialNo = string.Empty;			//일련번호
			string BPLID = string.Empty;			//사업장
			string RspCode = string.Empty;			//담당
			string CntcCode = string.Empty;			//사번
			string CntcName = string.Empty;			//성명
			string CardCode = string.Empty;			//거래처
			string CardName = string.Empty;			//거래처명
			string DocDate = string.Empty;			//발행일자
			decimal Amount = 0;			            //공급가액
			decimal VatTax = 0;			            //부가가치세
			string Contents = string.Empty;			//내용
			string UserSign = string.Empty;         //UserSign

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				SerialNo = oForm.Items.Item("SerialNo").Specific.Value.ToString().Trim();
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				CardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				Amount = Convert.ToDecimal(oForm.Items.Item("Amount").Specific.Value.ToString().Trim());
				VatTax = Convert.ToDecimal(oForm.Items.Item("VatTax").Specific.Value.ToString().Trim());
				Contents = oForm.Items.Item("Contents").Specific.Value.ToString().Trim();
				UserSign = PSH_Globals.oCompany.UserSignature.ToString();

				// DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
				sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_FI421H]";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					DocEntry = 1;
				}
				else
				{
					DocEntry = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1;
				}

				sQry = "EXEC [PS_FI421_02] '";
				sQry += DocEntry + "','";
				sQry += SerialNo + "','";
				sQry += BPLID + "','";
				sQry += RspCode + "','";
				sQry += CntcCode + "','";
				sQry += CntcName + "','";
				sQry += CardCode + "','";
				sQry += CardName + "','";
				sQry += DocDate + "','";
				sQry += Amount + "','";
				sQry += VatTax + "','";
				sQry += Contents + "','";
				sQry += UserSign + "'";

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
		/// PS_FI421_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_FI421_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;

			int ErrNum = 0;

			try
			{
				if (oForm.Items.Item("RspCode").Specific.Value.ToString().Trim() == "%")                    //담당
				{
					ErrNum = 1;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))    //거래처
				{   ErrNum = 2;
					throw new Exception();
				}
				if (oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() == "%")                    //발행일자
				{
					ErrNum = 3;
					throw new Exception();
				}
				if (oForm.Items.Item("Amount").Specific.Value.ToString().Trim() == "0")                     //공급가액
				{
					ErrNum = 4;
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("Contents").Specific.Value.ToString().Trim()))    //내용
				{
					ErrNum = 5;
					throw new Exception();
				}
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("담당은 필수사항입니다. 선택하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 2)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("거래처는 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 3)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("발행일자는 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 4)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("공급가액은 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else if (ErrNum == 5)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("내용은 필수사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}

			return functionReturnValue;
		}

		/// <summary>
		/// PS_FI421_CreateItems
		/// </summary>
		/// <returns></returns>
		private void PS_FI421_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oDS_PS_FI421H = oForm.DataSources.DBDataSources.Item("@PS_FI421H");
				oDS_PS_FI421L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				// 메트릭스 개체 할당
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				//관리번호
				oForm.DataSources.UserDataSources.Add("SDocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("SDocEntry").Specific.DataBind.SetBound(true, "", "SDocEntry");

				//일련번호
				oForm.DataSources.UserDataSources.Add("SSerialNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SSerialNo").Specific.DataBind.SetBound(true, "", "SSerialNo");

				//사업장_S
				oForm.DataSources.UserDataSources.Add("SBPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("SBPLId").Specific.DataBind.SetBound(true, "", "SBPLId");
				//사업장_E

				//담당
				oForm.DataSources.UserDataSources.Add("SRspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("SRspCode").Specific.DataBind.SetBound(true, "", "SRspCode");

				//사번
				oForm.DataSources.UserDataSources.Add("SCntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SCntcCode").Specific.DataBind.SetBound(true, "", "SCntcCode");

				//성명
				oForm.DataSources.UserDataSources.Add("SCntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("SCntcName").Specific.DataBind.SetBound(true, "", "SCntcName");

				//거래처
				oForm.DataSources.UserDataSources.Add("SCardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SCardCode").Specific.DataBind.SetBound(true, "", "SCardCode");

				//거래처명
				oForm.DataSources.UserDataSources.Add("SCardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("SCardName").Specific.DataBind.SetBound(true, "", "SCardName");

				//발행일자FR
				oForm.DataSources.UserDataSources.Add("SDocDateFr", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SDocDateFr").Specific.DataBind.SetBound(true, "", "SDocDateFr");

				//발행일자TO
				oForm.DataSources.UserDataSources.Add("SDocDateTo", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("SDocDateTo").Specific.DataBind.SetBound(true, "", "SDocDateTo");

				//공급가액
				oForm.DataSources.UserDataSources.Add("SAmount", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("SAmount").Specific.DataBind.SetBound(true, "", "SAmount");

				//부가가치세
				oForm.DataSources.UserDataSources.Add("SVatTax", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("SVatTax").Specific.DataBind.SetBound(true, "", "SVatTax");

				//내용
				oForm.DataSources.UserDataSources.Add("SContents", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("SContents").Specific.DataBind.SetBound(true, "", "SContents");
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
		/// PS_FI421_ComboBox_Setting
		/// </summary>
		private void PS_FI421_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("SBPLId").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("SBPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//담당(기본정보)
				oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'F003'", "1", false, false);
				oForm.Items.Item("RspCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//담당(조회조건)
				oForm.Items.Item("SRspCode").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("SRspCode").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'F003'", "1", false, false);
				oForm.Items.Item("SRspCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//매트릭스
				//사업장
				dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");

				//담당
				dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("RspCode"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'F003'", "", "");
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
		/// PS_FI421_EnableMenus
		/// </summary>
		private void PS_FI421_EnableMenus()
		{
			try
			{
				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// FormResize
		/// </summary>
		private void PS_FI421_FormResize()
		{
			try
			{
				oMat01.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// FormReset
		/// </summary>
		private void PS_FI421_FormReset()
		{
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			try
			{
				oForm.Freeze(true);

				//관리번호
				sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_FI421H]";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					oDS_PS_FI421H.SetValue("DocEntry", 0, "1");
				}
				else
				{
					oDS_PS_FI421H.SetValue("DocEntry", 0, Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1));
				}

				oDS_PS_FI421H.SetValue("U_BPLId", 0, dataHelpClass.User_BPLID());               //사업장
				oDS_PS_FI421H.SetValue("U_RspCode", 0, "%");                //담당
				oDS_PS_FI421H.SetValue("U_CntcCode", 0, "");                //사번
				oDS_PS_FI421H.SetValue("U_CntcName", 0, "");                //성명
				oDS_PS_FI421H.SetValue("U_CardCode", 0, "");                //거래처코드
				oDS_PS_FI421H.SetValue("U_CardName", 0, "");                //거래처명
				oDS_PS_FI421H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));      //발행일자
				oDS_PS_FI421H.SetValue("U_Amount", 0, "0");             //공급가액
				oDS_PS_FI421H.SetValue("U_VatTax", 0, "0");             //부가가치세
				oDS_PS_FI421H.SetValue("U_Contents", 0, "");                //내용

				PS_FI421_GetSerialNo(); //일련번호
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
		/// PS_FI421_GetSerialNo
		/// </summary>
		private void PS_FI421_GetSerialNo()
		{
			string DocDate = string.Empty;
			string BPLID = string.Empty;
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 6);
				sQry = "EXEC PS_FI421_05 '" + BPLID + "', '" + DocDate + "'";
				oRecordSet.DoQuery(sQry);

				oForm.Items.Item("SerialNo").Specific.Value = oRecordSet.Fields.Item("SerialNo").Value;
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
		/// PS_FI421_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_FI421_Print_Report01()
		{
			string WinTitle = null;
			string ReportName = null;
			string DocEntry = null;
			string BPLID = null;

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				// 인자 MOVE , Trim 시키기..
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

				WinTitle = "[PS_FI421] 입금표";


				if (BPLID == "1")  //창원
				{
					ReportName = "PS_FI421_01.rpt";   // 없슴
				}
				else if (BPLID == "2")  //동래
				{
					ReportName = "PS_FI421_02.rpt";
				}
				else if (BPLID == "3")  //사상
				{
					ReportName = "PS_FI421_03.rpt";   // 없슴
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				// Formula

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);

			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_FormItemEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				switch (pVal.EventType)
				{
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:					//1
						Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:						//2
						Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:					//5
						Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:						    //6
						Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:					//7
						Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:			//8
						Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:						//10
						Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:					//11
						//Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:					//18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:				//19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:					//20
						Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:				//27
						//Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:						//3
						Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:						//4
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:					//17
						Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
					if (pVal.ItemUID == "BtnAdd")  // 추가/확인 버튼클릭
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_FI421_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_FI421_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_FI421_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_FI421_LoadCaption();
							PS_FI421_MTX01();
							oLast_Mode = Convert.ToInt16(oForm.Mode);
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_FI421_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_FI421_UpdateData() == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_FI421_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_FI421_LoadCaption();
							PS_FI421_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSearch")   //조회
					{
						PS_FI421_FormReset();
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_FI421_LoadCaption();
						PS_FI421_MTX01();
					}
					else if (pVal.ItemUID == "BtnDelete")  //삭제
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
						{
							PS_FI421_DeleteData();
							PS_FI421_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_FI421_LoadCaption();
							PS_FI421_MTX01();
						}
						else
						{
						}
					}
					else if (pVal.ItemUID == "BtnPrint")  //입금표출력
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_FI421_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");                  ////사용자값활성(사번)
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");                  ////사용자값활성(거래처)

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SCntcCode", "");                 ////사용자값활성(사번)
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SCardCode", "");					////사용자값활성(거래처)
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
							oForm.Freeze(true);

							oDS_PS_FI421H.SetValue("DocEntry", 0, oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value);						//관리번호
							oDS_PS_FI421H.SetValue("U_SerialNo", 0, oMat01.Columns.Item("SerialNo").Cells.Item(pVal.Row).Specific.Value);					//일련번호
							oDS_PS_FI421H.SetValue("U_BPLId", 0, oMat01.Columns.Item("BPLId").Cells.Item(pVal.Row).Specific.Value);							//사업장
							oDS_PS_FI421H.SetValue("U_RspCode", 0, oMat01.Columns.Item("RspCode").Cells.Item(pVal.Row).Specific.Value);						//담당
							oDS_PS_FI421H.SetValue("U_CntcCode", 0, oMat01.Columns.Item("CntcCode").Cells.Item(pVal.Row).Specific.Value);					//사번
							oDS_PS_FI421H.SetValue("U_CntcName", 0, oMat01.Columns.Item("CntcName").Cells.Item(pVal.Row).Specific.Value);					//성명
							oDS_PS_FI421H.SetValue("U_CardCode", 0, oMat01.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific.Value);					//거래처
							oDS_PS_FI421H.SetValue("U_CardName", 0, oMat01.Columns.Item("CardName").Cells.Item(pVal.Row).Specific.Value);					//거래처명
							oDS_PS_FI421H.SetValue("U_DocDate", 0, oMat01.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value);						//발행일자
							oDS_PS_FI421H.SetValue("U_Amount", 0, oMat01.Columns.Item("Amount").Cells.Item(pVal.Row).Specific.Value);						//공급가액
							oDS_PS_FI421H.SetValue("U_VatTax", 0, oMat01.Columns.Item("VatTax").Cells.Item(pVal.Row).Specific.Value);						//부가가치세
							oDS_PS_FI421H.SetValue("U_Contents", 0, oMat01.Columns.Item("Contents").Cells.Item(pVal.Row).Specific.Value);					//내용

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_FI421_LoadCaption();
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_EVENT_MATRIX_LINK_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
						if (pVal.ItemUID == "Mat01")
						{
						}
						else
						{
							if (pVal.ItemUID == "CntcCode")
							{
								oForm.Items.Item("CntcName").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.Value + "'", "");								//성명
							}
							else if (pVal.ItemUID == "SCntcCode")
							{
								oForm.Items.Item("SCntcName").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("SCntcCode").Specific.Value + "'", "");
							}
							else if (pVal.ItemUID == "CardCode")
							{
								oForm.Items.Item("CardName").Specific.Value = dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item("CardCode").Specific.Value + "'", "");
							}
							else if (pVal.ItemUID == "SCardCode")
							{
								oForm.Items.Item("SCardName").Specific.Value = dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item("SCardCode").Specific.Value + "'", "");
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal,ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_FI421_FormResize();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
			finally
			{
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
					SubMain.Remove_Forms(oFormUniqueID01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FI421H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FI421L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
			int i = 0;
			try
			{
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
					}
					else if (pVal.BeforeAction == false)
					{
						for (i = 1; i <= oMat01.VisualRowCount; i++)
						{
							oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat01.FlushToDataSource();
						oDS_PS_FI421H.RemoveRecord(oDS_PS_FI421H.Size - 1);
						oMat01.LoadFromDataSource();
						if (oMat01.RowCount == 0)
						{
							PS_FI421_Add_MatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_FI421H.GetValue("U_CntcCode", oMat01.RowCount - 1).ToString().Trim()))
							{
								PS_FI421_Add_MatrixRow(oMat01.RowCount, false);
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
			}
		}

		/// <summary>
		/// Raise_FormMenuEvent
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
						case "1284":                            //취소
							break;
						case "1286":                            //닫기
							break;
						case "1293":                            //행삭제
							break;
						case "1281":                            //찾기
							break;
						case "1282":                            //추가
							// 추가버튼 클릭시 메트릭스 insertrow									
							PS_FI421_FormReset();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_FI421_LoadCaption();
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                            //레코드이동버튼
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284":                            //취소
							break;
						case "1286":                            //닫기
							break;
						case "1293":                            //행삭제
							break;
						case "1281":                            //찾기
							break;
						case "1282":                            //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                            //레코드이동버튼
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
			}
		}

		/// <summary>
		/// Raise_FormDataEvent
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                         //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                          //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                       //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                       //36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                         //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                          //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                       //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                       //36
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
			finally
			{
			}
		}
	}
}
