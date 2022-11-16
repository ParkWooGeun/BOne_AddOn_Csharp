using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn.Core
{
	/// <summary>
	/// 판매오더
	/// </summary>
	internal class S139 : PSH_BaseClass
	{
		private SAPbouiCOM.Matrix oMat01;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private bool oSetBackOrderFunction01;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="formUID"></param>
		public override void LoadForm(string formUID)
		{
			try
			{
				oForm = PSH_Globals.SBO_Application.Forms.Item(formUID);
				oForm.Freeze(true);
				oMat01 = oForm.Items.Item("38").Specific;
				SubMain.Add_Forms(this, formUID, "S139");

                S139_CreateItems();
                S139_EnableFormItem(false);
            }
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void S139_CreateItems()
        {
            SAPbouiCOM.Item oNewITEM = null;
            
            try
            {
                oNewITEM = oForm.Items.Add("TradeType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oNewITEM.Left = oForm.Items.Item("2003").Left;
                oNewITEM.Top = (oForm.Items.Item("2003").Top + oForm.Items.Item("2003").Height) + 1;
                oNewITEM.Height = oForm.Items.Item("2003").Height;
                oNewITEM.Width = oForm.Items.Item("2003").Width;
                oNewITEM.DisplayDesc = true;
                oNewITEM.Specific.DataBind.SetBound(true, "ORDR", "U_TradeType");
                oNewITEM.Specific.ValidValues.Add("1", "일반");
                oNewITEM.Specific.ValidValues.Add("2", "임가공");
                oNewITEM.Specific.ValidValues.Add("3", "선생산");

                oNewITEM = oForm.Items.Add("Static01", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("2002").Left;
                oNewITEM.Top = (oForm.Items.Item("2002").Top + oForm.Items.Item("2002").Height) + 1;
                oNewITEM.Height = oForm.Items.Item("2002").Height;
                oNewITEM.Width = oForm.Items.Item("2002").Width;
                oNewITEM.LinkTo = "TradeType";
                oNewITEM.Specific.Caption = "거래형태";

                oNewITEM = oForm.Items.Add("DCardCod", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oNewITEM.Left = oForm.Items.Item("222").Left;
                oNewITEM.Top = (oForm.Items.Item("222").Top + oForm.Items.Item("222").Height) + 1;
                oNewITEM.Height = oForm.Items.Item("222").Height;
                oNewITEM.Width = oForm.Items.Item("222").Width;
                oNewITEM.Specific.DataBind.SetBound(true, "ORDR", "U_DCardCod");

                oNewITEM = oForm.Items.Add("Static03", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("230").Left;
                oNewITEM.Top = (oForm.Items.Item("230").Top + oForm.Items.Item("230").Height) + 1;
                oNewITEM.Height = oForm.Items.Item("230").Height;
                oNewITEM.Width = oForm.Items.Item("230").Width;
                oNewITEM.LinkTo = "DCardCod";
                oNewITEM.Specific.Caption = "납품처코드";

                oNewITEM = oForm.Items.Add("DCardNam", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oNewITEM.Left = oForm.Items.Item("DCardCod").Left;
                oNewITEM.Top = (oForm.Items.Item("DCardCod").Top + oForm.Items.Item("DCardCod").Height) + 1;
                oNewITEM.Height = oForm.Items.Item("DCardCod").Height;
                oNewITEM.Width = oForm.Items.Item("DCardCod").Width;
                oNewITEM.Enabled = false;
                oNewITEM.Specific.DataBind.SetBound(true, "ORDR", "U_DCardNam");

                oNewITEM = oForm.Items.Add("Static04", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("Static03").Left;
                oNewITEM.Top = (oForm.Items.Item("Static03").Top + oForm.Items.Item("Static03").Height) + 1;
                oNewITEM.Height = oForm.Items.Item("Static03").Height;
                oNewITEM.Width = oForm.Items.Item("Static03").Width;
                oNewITEM.LinkTo = "DCardNam";
                oNewITEM.Specific.Caption = "납품처명";

                oNewITEM = oForm.Items.Add("LotNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oNewITEM.Left = oForm.Items.Item("DCardNam").Left;
                oNewITEM.Top = (oForm.Items.Item("DCardNam").Top + oForm.Items.Item("DCardNam").Height) + 1;
                oNewITEM.Height = oForm.Items.Item("DCardNam").Height;
                oNewITEM.Width = oForm.Items.Item("DCardNam").Width;
                oNewITEM.Specific.DataBind.SetBound(true, "ORDR", "U_LotNo");

                oNewITEM = oForm.Items.Add("Static05", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Left = oForm.Items.Item("Static04").Left;
                oNewITEM.Top = (oForm.Items.Item("Static04").Top + oForm.Items.Item("Static04").Height) + 1;
                oNewITEM.Height = oForm.Items.Item("Static04").Height;
                oNewITEM.Width = oForm.Items.Item("Static04").Width;
                oNewITEM.LinkTo = "LotNo";
                oNewITEM.Specific.Caption = "업체수주번호";

                oNewITEM = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oNewITEM.Top = oForm.Items.Item("1").Top - 12;
                oNewITEM.Left = oForm.Items.Item("1").Left;
                oNewITEM.Height = 12;
                oNewITEM.Width = 120;
                oNewITEM.FontSize = 10;
                oNewITEM.Specific.Caption = "Addon running";
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oNewITEM);
            }
        }

        /// <summary>
        /// 각 모드에 따른 아이템설정
        /// </summary>
        /// <param name="Status"></param>
        private void S139_EnableFormItem(bool Status)
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("TradeType").Enabled = true;
                    oForm.Items.Item("DCardNam").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("TradeType").Enabled = true;
                    oForm.Items.Item("DCardNam").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    if (oForm.Items.Item("2001").Specific.Value.ToString().Trim() == "2" 
                        && oForm.Items.Item("81").Specific.Value.ToString().Trim() != "3" 
                        && oForm.Items.Item("81").Specific.Value.ToString().Trim() != "4") //동래공장인 경우 수정가능토록 취소, 종료가 아닐경우
                    {
                        
                        oForm.Items.Item("TradeType").Enabled = true;
                    }
                    else
                    {
                        oForm.Items.Item("TradeType").Enabled = false;
                    }

                    oForm.Items.Item("DCardNam").Enabled = false;
                }

                if (Status == true)
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        if (oForm.Items.Item("2001").Specific.Value.ToString().Trim() == "2"
                            && oForm.Items.Item("81").Specific.Value.ToString().Trim() != "3"
                            && oForm.Items.Item("81").Specific.Value.ToString().Trim() != "4") //동래공장인 경우 수정가능토록 취소, 종료가 아닐경우
                        {
                            oForm.Items.Item("TradeType").Enabled = true;
                        }
                        else
                        {
                            oForm.Items.Item("TradeType").Enabled = false;
                        }
                        oForm.Items.Item("DCardNam").Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <param name="FormMode"></param>
        /// <returns></returns>
        private bool S139_CheckDataValid(SAPbouiCOM.BoFormMode FormMode)
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            string Query01;
            string Query02;
            string CardCode;
            string DocDate;
            string ItemCode;
            string Text;
            string SAmt;
            double Amt;
            double OverAmt;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("4").Specific.Value))
                {
                    errMessage = "고객은 필수입니다.";
                    oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("2001").Specific.Value.ToString().Trim()))
                {
                    errMessage = "사업장은 필수입니다.";
                    oForm.Items.Item("2001").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("TradeType").Specific.Value))
                {
                    errMessage = "거래형태는 필수입니다.";
                    oForm.Items.Item("TradeType").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }
                else if (oForm.Items.Item("2001").Specific.Value.ToString().Trim() != "1" && oForm.Items.Item("TradeType").Specific.Selected.Value == "2") //창원이 아닌데 임가공이 선택된 경우
                {
                    errMessage = "창원사업장이 아닌경우 임가공거래가 불가능합니다.";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount <= 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                CardCode = oForm.Items.Item("4").Specific.Value;
                DocDate = oForm.Items.Item("10").Specific.Value;
                Text = oForm.Items.Item("29").Specific.Value;
                SAmt = Text.Split(' ')[0]; //수주총계, 통화 단위 제거
                Amt = Convert.ToDouble(SAmt.Replace(",", ""));
                ItemCode = oMat01.Columns.Item("1").Cells.Item(1).Specific.Value;

                Query01 = "Select U_ItmBsort From OITM Where ItemCode = '" + ItemCode + "'";

                RecordSet01.DoQuery(Query01);
                
                if (RecordSet01.Fields.Item(0).Value == "111") //분말
                {
                    Query02 = "EXEC [S139_hando] '" + CardCode + "', '" + DocDate + "'";
                    RecordSet02.DoQuery(Query02);

                    OverAmt = Convert.ToDouble(RecordSet02.Fields.Item("OverAmt").Value);

                    if (OverAmt - Amt < 0)
                    {
                        errMessage = "여신한도를 초과합니다. 확인바랍니다.";
                        throw new Exception();
                    }
                }

                for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품목은 필수입니다.";
                        oMat01.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (Convert.ToDouble(oMat01.Columns.Item("11").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "수량(중량)은 필수입니다.";
                        oMat01.Columns.Item("11").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (string.IsNullOrEmpty(oMat01.Columns.Item("14").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "단가는 필수입니다.";
                        oMat01.Columns.Item("14").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    
                    if (oForm.Items.Item("70").Specific.Selected.Value == "S" || oForm.Items.Item("70").Specific.Selected.Value == "L") //현지, 시스템통화
                    {
                        if (codeHelpClass.Right(oMat01.Columns.Item("14").Cells.Item(i).Specific.Value, 3) != "KRW")
                        {
                            errMessage = "헤더와 라인의 통화가 다릅니다.";
                            throw new Exception();
                        }
                    }
                    
                    if (oForm.Items.Item("70").Specific.Selected.Value == "C") //BP통화
                    {
                        if (oForm.Items.Item("63").Specific.Value != codeHelpClass.Right(oMat01.Columns.Item("14").Cells.Item(i).Specific.Value, 3)) //DocCur와 Price의 마지막3자리 비교
                        {
                            errMessage = "헤더와 라인의 통화가 다릅니다.";
                            throw new Exception();
                        }
                    }
                    
                    if (oForm.Items.Item("TradeType").Specific.Selected.Value == "1") //일반
                    {
                        if (dataHelpClass.GetItem_TradeType(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value) == "2") //품목 : 임가공
                        {
                            errMessage = "문서의 거래형태와 품목의 거래형태가 다릅니다.";
                            oMat01.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                    
                    if (oForm.Items.Item("TradeType").Specific.Selected.Value == "2") //임가공
                    {
                        if (dataHelpClass.GetItem_TradeType(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value) == "1") //품목 : 일반
                        {
                            errMessage = "문서의 거래형태와 품목의 거래형태가 다릅니다.";
                            oMat01.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                }

                if (FormMode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    if (S139_CheckValidate("검사") == false)
                    {
                        returnValue = false;
                        return returnValue;
                    }
                }

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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool S139_CheckValidate(string ValidateType)
        {
            string errMessage = string.Empty;
            bool returnValue = false;
            string Query01;
            bool Exist;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (ValidateType == "검사")
                {
                    //출하요청등록된것이 존재하면
                    if (Convert.ToInt16(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" + oForm.Items.Item("8").Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                    {
                        if (oForm.Items.Item("70").Specific.Selected.Value.ToString().Trim() != dataHelpClass.GetValue("SELECT CurSource FROM [ORDR] WHERE DocEntry = '" + oForm.Items.Item("8").Specific.Value.ToString().Trim() + "'", 0, 1)) //폼의 현지통화와, DB의 현지통화가 다르면
                        {
                            errMessage = "출하,선출요청된 문서입니다. 통화가 변경되었습니다.";
                            throw new Exception();
                        }

                        if (oForm.Items.Item("63").Specific.Selected.Value.ToString().Trim() != dataHelpClass.GetValue("SELECT DocCur FROM [ORDR] WHERE DocEntry = '" + oForm.Items.Item("8").Specific.Value.ToString().Trim() + "'", 0, 1)) //폼의 현지통화와, DB의 현지통화가 다르면
                        {
                            errMessage = "출하,선출요청된 문서입니다. 통화가 변경되었습니다.";
                            throw new Exception();
                        }

                        if (oForm.Items.Item("2001").Specific.Selected.Value.ToString().Trim() != dataHelpClass.GetValue("SELECT BPLId FROM [ORDR] WHERE DocEntry = '" + oForm.Items.Item("8").Specific.Value.ToString().Trim() + "'", 0, 1)) //폼의 사업장과, DB의 사업장이 다르면
                        {
                            errMessage = "출하,선출요청된 문서입니다. 사업장이 변경되었습니다.";
                            throw new Exception();
                        }
                    }

                    //라인
                    Exist = false;
                    Query01 = "SELECT DocEntry,LineNum FROM [RDR1] WHERE DocEntry = '" + oForm.Items.Item("8").Specific.Value + "'";
                    RecordSet01.DoQuery(Query01);

                    for (int i = 0; i <= RecordSet01.RecordCount - 1; i++)
                    {
                        Exist = false;

                        for (int j = 1; j <= oMat01.RowCount - 1; j++)
                        {
                            //라인번호가 같고, 품목코드가 같으면 존재하는 행 , LineNum에 값이 존재하는지 확인필요(행삭제된 행인 경우 LineNum이 존재하지않음)
                            string tempLineNum = oMat01.Columns.Item("U_LineNum").Cells.Item(j).Specific.Value == "" ? "-1" : oMat01.Columns.Item("U_LineNum").Cells.Item(j).Specific.Value;

                            if (Convert.ToInt16(RecordSet01.Fields.Item(1).Value) == Convert.ToInt16(tempLineNum))
                            {
                                Exist = true;
                            }
                        }
                        
                        if (Exist == false) //삭제된 행중
                        {
                            if (Convert.ToInt16(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" + RecordSet01.Fields.Item(0).Value.ToString().Trim() + "' AND PS_SD030L.U_RDR1Num = '" + RecordSet01.Fields.Item(1).Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errMessage = "삭제된 행이 다른사용자에 의해 출하,선출요청되었습니다. 적용할 수 없습니다.";
                                throw new Exception();
                            }
                        }

                        RecordSet01.MoveNext();
                    }
                    
                    for (int i = 1; i <= oMat01.VisualRowCount - 1; i++) //수량가능성검사
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("U_LineNum").Cells.Item(i).Specific.Value))
                        {
                            //새로추가된 행인경우, 검사 불필요
                        }
                        else
                        {
                            //매트릭스에 입력된 수량과 DB상에 존재하는 수량의 값비교
                            if (Convert.ToDouble(oMat01.Columns.Item("11").Cells.Item(i).Specific.Value.ToString().Trim()) < Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(U_Weight) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" + oForm.Items.Item("8").Specific.Value.ToString().Trim() + "' AND PS_SD030L.U_RDR1Num = '" + oMat01.Columns.Item("U_LineNum").Cells.Item(i).Specific.Value.ToString().Trim() + "'", 0, 1)))
                            {
                                errMessage = i + "행의 수량이 출하요청,선출요청 수량보다 작습니다.";
                                throw new Exception();
                            }

                            //이미,출하선출된 행이 있으면 값이 수정불가
                            if (Convert.ToInt16(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" + oForm.Items.Item("8").Specific.Value.ToString().Trim() + "' AND PS_SD030L.U_RDR1Num = '" + oMat01.Columns.Item("U_LineNum").Cells.Item(i).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                //품목코드는 변경되면 안된다.

                                Query01 = "  SELECT     ItemCode, ";
                                Query01 += "            Price, ";
                                Query01 += "            U_TrType ";
                                Query01 += " FROM       [RDR1] RDR1";
                                Query01 += " WHERE      DocEntry = '" + oForm.Items.Item("8").Specific.Value.ToString().Trim() + "'";
                                Query01 += "            AND LineNum = '" + oMat01.Columns.Item("U_LineNum").Cells.Item(i).Specific.Value.ToString().Trim() + "'";

                                RecordSet01.DoQuery(Query01);
                                if (RecordSet01.Fields.Item(0).Value.ToString().Trim() == oMat01.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()
                                    && Convert.ToDouble(RecordSet01.Fields.Item(1).Value) == Convert.ToDouble(oMat01.Columns.Item("14").Cells.Item(i).Specific.Value.ToString().Split(" ")[0]) //단가
                                    && RecordSet01.Fields.Item(2).Value.ToString().Trim() == oMat01.Columns.Item("U_TrType").Cells.Item(i).Specific.Selected.Value.ToString().Trim())
                                {
                                }
                                else
                                {
                                    errMessage = "이미출하,선출요청된 행입니다. 수정할 수 없습니다.";
                                    throw new Exception();
                                }
                            }
                        }
                    }
                }
                else if (ValidateType == "수정")
                {
                }
                else if (ValidateType == "행삭제")
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) //추가,수정모드일때행삭제가능검사
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("U_LineNum").Cells.Item(oLastColRow01).Specific.Value))
                        {
                            //새로추가된 행인경우, 삭제 가능
                        }
                        else
                        {
                            if (Convert.ToInt16(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" + oForm.Items.Item("8").Specific.Value.ToString().Trim() + "' AND PS_SD030L.U_RDR1Num = '" + oMat01.Columns.Item("U_LineNum").Cells.Item(oLastColRow01).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errMessage = "이미출하,선출요청된 행입니다. 삭제할 수 없습니다.";
                                throw new Exception();
                            }

                            //작업지시가 존재하면 삭제 불가
                            if (Convert.ToInt16(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] WHERE Canceled = 'N' AND U_ItemCode = '" + oMat01.Columns.Item("1").Cells.Item(oLastColRow01).Specific.Value + "'", 0, 1)) > 0)
                            {
                                errMessage = "해당 작번은 작업지시가 존재합니다. 삭제할 수 없습니다.";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "취소")
                {
                    Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" + oForm.Items.Item("8").Specific.Value + "'";
                    RecordSet01.DoQuery(Query01);
                    for (int i = 0; i <= RecordSet01.RecordCount - 1; i++)
                    {
                        if (Convert.ToInt16(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND PS_SD030L.U_ORDRNum = '" + RecordSet01.Fields.Item(0).Value.ToString().Trim() + "' AND PS_SD030L.U_RDR1Num = '" + RecordSet01.Fields.Item(1).Value.ToString().Trim() + "'", 0, 1)) > 0)
                        {
                            errMessage = "출하,선출요청된문서입니다. 적용할 수 없습니다.";
                            throw new Exception();
                        }

                        RecordSet01.MoveNext();
                    }
                }
                else if (ValidateType == "닫기")
                {
                    Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" + oForm.Items.Item("8").Specific.Value + "'";
                    RecordSet01.DoQuery(Query01);
                    for (int i = 0; i <= RecordSet01.RecordCount - 1; i++)
                    {
                        if (Convert.ToInt16(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Status = 'O' AND PS_SD030L.U_ORDRNum = '" + RecordSet01.Fields.Item(0).Value.ToString().Trim() + "' AND PS_SD030L.U_RDR1Num = '" + RecordSet01.Fields.Item(1).Value.ToString().Trim() + "'", 0, 1)) > 0)
                        {
                            errMessage = "출하,선출요청된문서입니다. 적용할 수 없습니다.";
                            throw new Exception();
                        }

                        //작업지시가 존재하면 닫기 불가
                        if (Convert.ToInt16(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] WHERE Status = 'O' AND U_ItemCode = '" + oMat01.Columns.Item("1").Cells.Item(i + 1).Specific.Value.ToString().Trim() + "' AND U_SjNum = '" + RecordSet01.Fields.Item(0).Value.ToString().Trim() + "' AND U_SjLine = '" + RecordSet01.Fields.Item(1).Value.ToString().Trim() + "'", 0, 1)) > 0)
                        {
                            errMessage = i + 1 + "행의 작번이 작업지시가 존재합니다. 문서를 닫기(종료) 처리할 수 없습니다.";
                            throw new Exception();
                        }

                        RecordSet01.MoveNext();
                    }
                }

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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 기계사업부 작번닫기
        /// </summary>
        /// <param name="p_Row"></param>
        private void S139_CancelORDR(int p_Row)
        {
            //short l_ErrNum = 0;
            string errMessage = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet03 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string l_DocEntry; //문서번호
            string l_CardCode; //거래처코드
            string l_ItemCode; //품목코드(작번)
            string l_WOEntry; //작업지시문서번호
            
            try
            {
                l_DocEntry = oForm.Items.Item("8").Specific.Value;
                l_CardCode = oForm.Items.Item("4").Specific.Value;
                l_ItemCode = oMat01.Columns.Item("1").Cells.Item(p_Row).Specific.Value;

                //이미 등록된 작번 닫기(종료) 건인지 검사_S
                sQry = "  SELECT    'X' ";
                sQry += " FROM      Z_ORDR_Cancel ";
                sQry += " WHERE     DocEntry = " + l_DocEntry;
                sQry += "           AND CardCode = '" + l_CardCode + "'";
                sQry += "           AND ItemCode = '" + l_ItemCode + "'";

                oRecordSet02.DoQuery(sQry);

                if (oRecordSet02.RecordCount != 0)
                {
                    errMessage = "이미 작번 닫기(종료) 처리가 등록된 작번입니다. 다시 확인 하십시오.";
                    throw new Exception();
                }
                //이미 등록된 작번 닫기(종료) 건인지 검사_E

                //작업지시가 존재하는지 검사_S
                sQry = "  SELECT    DocEntry ";
                sQry += " FROM      [@PS_PP030H] ";
                sQry += " WHERE     U_ItemCode = '" + l_ItemCode + "'";
                sQry += "           AND U_SjNum = '" + l_DocEntry + "'";
                sQry += "           AND Status = 'O'";
                sQry += "           AND U_OrdSub1 = '00'";
                sQry += "           AND U_OrdSub2 = '000'";

                oRecordSet03.DoQuery(sQry);

                if (oRecordSet03.RecordCount != 0)
                {
                    l_WOEntry = oRecordSet03.Fields.Item(0).Value;
                    errMessage = "작업지시가 존재하는 작번입니다. 작업지시를 먼저 닫기(종료)하십시오. " + (char)13 + "작업지시문서번호 : " + l_WOEntry;
                    throw new Exception();
                }
                //작업지시가 존재하는지 검사_E

                //최종 실행
                if (PSH_Globals.SBO_Application.MessageBox("작번 : " + l_ItemCode + "를(을) 닫기(종료) 처리하시겠습니까?", 1, "예", "아니오") == 1)
                {
                    sQry = "INSERT INTO Z_ORDR_Cancel VALUES (" + l_DocEntry + ",'" + l_CardCode + "','" + l_ItemCode + "', GETDATE(), " + PSH_Globals.oCompany.UserSignature + ")";
                    oRecordSet01.DoQuery(sQry);
                    PSH_Globals.SBO_Application.StatusBar.SetText("작번 닫기(종료) 처리가 완료되었습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet03);
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
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (S139_CheckDataValid(oForm.Mode) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (S139_CheckDataValid(oForm.Mode) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
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
                                S139_EnableFormItem(false);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                S139_EnableFormItem(false);
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
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "DCardCod", "");

                    if (pVal.ItemUID == "38")
                    {
                        if (S139_CheckValidate("수정") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        
                        if (pVal.ColUID == "1") //품목코드
                        {
                            if (pVal.CharPressed == 9)
                            {
                                PS_SM020 tempForm = new PS_SM020();
                                tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, oMat01.VisualRowCount, oForm.Items.Item("TradeType").Specific.Selected.Value.ToString().Trim());
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.ItemUID == "38")
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

                if (pVal.BeforeAction == true)
                {

                }
                else if (pVal.BeforeAction == false)
                {
                    if (oSetBackOrderFunction01 == true)
                    {
                        oSetBackOrderFunction01 = false;
                        dataHelpClass.SBO_SetBackOrderFunction(ref oForm);
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "38")
                    {
                        if (S139_CheckValidate("수정") == false)
                        {
                            BubbleEvent = false;
                            oForm.Freeze(false);
                            return;
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    PSH_Globals.SBO_Application.Forms.GetForm(oForm.Type.ToString(), oForm.TypeCount).Update();
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        S139_EnableFormItem(true);
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "38")
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
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "10000330")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                oSetBackOrderFunction01 = true;
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string itemCode;
            string Query01;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "38") //매트릭스
                        {
                            itemCode = oMat01.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value;
                            
                            if (pVal.ColUID == "U_Qty") //수량
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value = 0; //수량
                                    oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1; //중량
                                }
                                else
                                {
                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {
                                        oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
                                        
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102") //EAUOM
                                    {
                                        if (Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode)) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode));
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        if ((Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        if (System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                            }
                            else if (pVal.ColUID == "11")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value = 0; //수량
                                    oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1; //중량
                                }
                                else
                                {
                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {   
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102") //EAUOM
                                    {   
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        if ((Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        if (System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0) == 0)
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                            }
                            else if (pVal.ColUID == "1")
                            {
                                if (oMat01.VisualRowCount > 1)
                                {
                                    if (oForm.Items.Item("2001").Specific.Value.ToString().Trim() == "2" 
                                        && oForm.Items.Item("81").Specific.Value.ToString().Trim() != "3" 
                                        && oForm.Items.Item("81").Specific.Value.ToString().Trim() != "4") //동래공장이면서 문서상태가 취소,종료가 아닐경우에 작번을 변경하려고 할 때만 체크
                                    {
                                        //저장하려는 작번으로 등록된 작업지시가 있는지 체크_S
                                        Query01 = "  SELECT     COUNT(*)";
                                        Query01 += " FROM       [@PS_PP030H]";
                                        Query01 += " WHERE      U_ItemCode = '" + oMat01.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value + "'";

                                        oRecordSet01.DoQuery(Query01);

                                        if (Convert.ToInt16(oRecordSet01.Fields.Item(0).Value) > 0)
                                        {
                                            PSH_Globals.SBO_Application.StatusBar.SetText("이미 작업지시가 등록된 작번입니다. 수정할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                            oMat01.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value = ""; //품목코드 필드에 빈값 강제 입력
                                        }
                                        //저장하려는 작번으로 등록된 작업지시가 있는지 체크_E
                                        
                                        oForm.Items.Item("TradeType").Enabled = true;
                                    }
                                    else
                                    {
                                        oForm.Items.Item("TradeType").Enabled = false;
                                    }
                                }
                                else
                                {
                                    oForm.Items.Item("TradeType").Enabled = true;
                                }
                            }
                        }
                        else if (pVal.ItemUID == "DCardCod")
                        {
                            oForm.Items.Item("DCardNam").Specific.Value = dataHelpClass.GetValue("SELECT CardName FROM OCRD WHERE CardCode = '" + oForm.Items.Item("DCardCod").Specific.Value + "'", 0, 1);
                        }
                    }

                    PSH_Globals.SBO_Application.Forms.GetForm(oForm.Type.ToString(), oForm.TypeCount).Update();
                }
                else if (pVal.Before_Action == false)
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        S139_EnableFormItem(true);
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oMat01.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
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
        /// FORM_ACTIVATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_ACTIVATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (oSetBackOrderFunction01 == true)
                    {
                        oSetBackOrderFunction01 = false;
                        dataHelpClass.SBO_SetBackOrderFunction(ref oForm);
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
        /// 행삭제 체크 메서드(Raise_FormMenuEvent 에서 사용)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (S139_CheckValidate("행삭제") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (oMat01.VisualRowCount > 1)
                        {
                            oForm.Items.Item("TradeType").Enabled = false;
                        }
                        else
                        {
                            oForm.Items.Item("TradeType").Enabled = true;
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
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            if (S139_CheckValidate("취소") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286": //닫기
                            if (S139_CheckValidate("닫기") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
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
                            break;
                        case "CancelORDR": //작번 닫기(종료)
                            S139_CancelORDR(oLastColRow01);
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
                            S139_EnableFormItem(false);
                            break;
                        case "1282": //추가
                            S139_EnableFormItem(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            oMat01.AutoResizeColumns();
                            S139_EnableFormItem(false);
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
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
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            SAPbouiCOM.MenuCreationParams oCreationPackage = null;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    ////C#Migration 완료 후 주석 해제 필요_S(PS_MM005 클래스 참조)
                    ////작번 닫기(종료) 생성
                    //if (pVal.ItemUID == "38" & pVal.Row > 0 && pVal.Row <= oMat01.RowCount)
                    //{
                    //    oCreationPackage = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    //    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    //    oCreationPackage.UniqueID = "CancelORDR";
                    //    oCreationPackage.String = "작번 닫기(종료)";
                    //    oCreationPackage.Enabled = true;

                    //    PSH_Globals.SBO_Application.Menus.Item("1280").SubMenus.AddEx(oCreationPackage);
                    //}
                    ////C#Migration 완료 후 주석 해제 필요_E(PS_MM005 클래스 참조)
                }
                else if (pVal.BeforeAction == false)
                {
                    ////C#Migration 완료 후 주석 해제 필요_S(PS_MM005 클래스 참조)
                    ////작번 닫기(종료) 삭제
                    //if (pVal.ItemUID == "38" && pVal.Row > 0)
                    //{
                    //    if (oMat01.RowCount >= pVal.Row)
                    //    {
                    //        PSH_Globals.SBO_Application.Menus.RemoveEx("CancelORDR");
                    //    }
                    //}
                    ////C#Migration 완료 후 주석 해제 필요_E(PS_MM005 클래스 참조)
                }

                if (pVal.ItemUID == "38")
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
                if (oCreationPackage != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreationPackage);
                }
            }
        }
    }
}
