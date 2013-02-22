			// Displays an OpenFileDialog so the user can select a Cursor.
			//OpenFileDialog ^ openFileDialog1 = gcnew OpenFileDialog();
			//openFileDialog1->Filter = "Universal Data Link *.udl";
			//openFileDialog1->Title = "Select a Cursor File";
			//openFileDialog1->FileName = "C:\\Users\\Travis.Powell\\Desktop\\test.udl";

			// Show the Dialog.
			// If the user clicked OK in the dialog and
			// a .CUR file was selected, open it.
			/*if (openFileDialog1->ShowDialog() == System::Windows::Forms::DialogResult::OK)
			{*/
				// Assign the cursor in the Stream to
				// the Form's Cursor property.
				//this->Cursor = gcnew System::Windows::Forms::Cursor(openFileDialog1->OpenFile());
				//System::IO::StreamReader ^ sr = gcnew System::IO::StreamReader(openFileDialog1->FileName);
				//MessageBox::Show(sr->ReadToEnd());
				//sr->Close();
			//}

			 /*String ^ constr = "File Name=C:\\Users\\Travis.Powell\\Desktop\\test.udl";
			 OleDbConnection ^ myconn = gcnew OleDbConnection(constr);
			 //myconn->Open();
			 //myconn->Close();

			 //MessageBox::Show(myconn->DataSource);
			 //MessageBox::Show(myconn->Provider);

			 String ^ strProvider = gcnew String("");
			 String ^ strDataSource = gcnew String("");

			 strProvider = myconn->Provider;
			 strDataSource = myconn->DataSource;
			 this->providerText->Text = System::String::Concat("Provider=", strProvider, ";Data Source=", strDataSource, ";Persist Security Info=False");
			 DSN = this->providerText->Text;
			 
			 /*myconn->ConnectionString::set(DSN);
			 myconn->Open();*/


			 /*System::Boolean connStatus = new System::Boolean();
			 connStatus = false;*/

			 //oledb::DataLinks dataLinks = gcnew oledb::DataLinksClass();

		//Spacer

			 /*p_conn->ConnectionString->set("Provider=OraOLEDB.Oracle.1;Password=travisp;Persist Security Info=True;User ID=tpowell;Data Source=TEST8");*/


		//Spacer

			 /*p_conn = p_IDSL->PromptNew();

			 connStatus = p_IDSL->PromptEdit(p_conn);*/

			 //if (p_conn->PromptEdit() == System::Windows::Forms::DialogResult::OK)
				 //MessageBox::Show(p_conn);

			 //connStatus = p_conn->
				//MessageBox::Show("The calculations are complete","My Application", MessageBoxButtons::OKCancel,MessageBoxIcon::Asterisk);

			 //DSN = gcnew String((char *)p_conn->ConnectionString);

			 //DSN = (*BSTRTest.m_Data).m_wstr;

			 //DSN = ADODB::_ConnectionPtr.get_ConnectionString(p_conn->ConnectionString);
			 //DSN = p_conn->ConnectionString;

			 /*if(!p_conn->ConnectionString){
				 this->providerText->Text = "Null";
			 }*/


			 //DSN = p_conn->ConnectionString;
			 //this->providerText->Text = DSN;

#include "stdafx.h"


					/*String^ fileName = "textfile.txt";
					System::IO::StreamWriter ^ sw = gcnew System::IO::StreamWriter(fileName);

					while (myReader->Read()) {
						info = "";
						if(printHeader == 1){
							for (int i = 0; myReader->FieldCount > i; i++){
								info += myReader->GetName(i) + "|";
							}
							printHeader = 0;
						}else{
							for (int i = 0; myReader->FieldCount > i; i++){
								info += myReader->GetValue(i) + "|";
							}
						}

						sw->WriteLine(info);
					}

						//sw->WriteLine(info);

					sw->Close();*/


				/*if(e->Errors->default[0]->NativeError == -524420377)		// The Microsoft Jet database engine cannot find the input...
					MessageBox::Show(e->Message,"Invalid Input", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				else if(e->Errors->default[0]->NativeError == -526650802)	// Syntax error in FROM clause.
					MessageBox::Show(e->Message,"Invalid Query", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				else if(e->Errors->default[0]->NativeError == -525995440)	// The SELECT statement includes a reserved word or an argument...
					MessageBox::Show(e->Message,"Invalid Input", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				else if(e->Errors->default[0]->NativeError == -533138860)	// Invalid SQL statement; expected 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT', or 'UPDATE'.
					MessageBox::Show(e->Message,"Invalid Query", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				else if(e->Errors->default[0]->NativeError == -524553244)	// Syntax error (missing operator) in query expression '* where'.
					MessageBox::Show(e->Message,"Invalid Query", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				else if(e->Errors->default[0]->NativeError == -525274553)	// Syntax error in WHERE clause.
					MessageBox::Show(e->Message,"Invalid Query", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);*/