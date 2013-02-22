/*   Evisions ADO Tool                                                                       */
/*                                                                                           */
/*   Developed by Travis Powell                                                              */
/*   Date Created: 1/25/2013                                                                 */
/*   Rev 1.24v                                                                               */
/*                                                                                           */
/*   ADO Tool Description                                                                    */
/*                                                                                           */
/*   This tool is used to connect to a Database outside the scope of any Evisions            */
/*   Applications.  This uses the Datalink Properties Tool found in the Microsoft Operating  */
/*   Systems.  It stores a temporary .txt file in the %USERPROFILE%\appdata\local\Temp       */
/*   filepath with the last generated connection string when using the Datalink Properties   */
/*   Window.                                                                                 */
/*                                                                                           */
/*                                                                                           */
/*   More R-Evisions to Come!                                                                */
/*                                                                                           */
/*                                                                                           */
/*                                                                                           */
/*   Rev 1.24v                                                                               */
/*   Modifications: 2/21/2013                                                                */
/*     - Fixed saving text to file for every keystroke                                       */
/*                                                                                           */
/*                                                                                           */
/*   Rev 1.23v                                                                               */
/*   Modifications: 2/19/2013                                                                */
/*     - Add a Row Count to the form                                                         */
/*     - Refresh connection string in temporary .txt file on the fly                         */
/*     - Clear button updated to only clear result set                                       */
/*     - Delete button, renamed to Clear All, clears and removes ALL data                    */
/*     - Add Evisions.com website link                                                       */
/*                                                                                           */
/*                                                                                           */
/*   Rev 1.22v                                                                               */
/*   Modifications: 2/19/2013                                                                */
/*     - Updated clearing out hidden variables when unchecking checkbox                      */
/*                                                                                           */
/*                                                                                           */
/*   Rev 1.21v                                                                               */
/*   Modifications: 2/18/2013                                                                */
/*     - Updated the password retrieval for Microsoft Access databases                       */
/*     - Updated comments for functions                                                      */
/*     - Added label on main page with Revision version                                      */
/*                                                                                           */
/*                                                                                           */
/*   Rev 1.20v                                                                               */
/*   Modifications: 2/16/2013                                                                */
/*   ENHANCEMENT Added                                                                       */
/*     - Blank username and password can be specified in the datalink properties window.     */
/*     - New form lets user put in username and password(masked) to be added to DSN prior    */
/*           to connecting to the database and running the query.                            */
/*                                                                                           */
/*                                                                                           */
/*   Rev 1.10v                                                                               */
/*   Modifications: 2/15/2013                                                                */
/*     - Exception Handling for OLEDB, File Read/Write,                                      */
/*     - Temporary File containing DSN to be used on Form.                                   */
/*     - Ability to Delete the temporary file.                                               */
/*     - Comments section for each function, what it does and its purpose.                   */
/*                                                                                           */

#import "c:\program files\common files\system\ole db\oledb32.dll" rename_namespace("oledb")
#import "C:\Program Files\Common Files\System\ado\msado15.dll" rename("EOF", "EOFile")

#include "frmPassword.h"

#pragma once

namespace ADO_tool {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::Data::OleDb;

	/// <summary>
	/// Summary for Form1
	/// </summary>
	public ref class Form1 : public System::Windows::Forms::Form
	{
	public:
		Form1(void)
		{
			InitializeComponent();

			/* This will check to see if the evi_ado.txt file that contains the connection string exists. */
			/* If it exists, it will retrieve the existing connection string and populate the Form.       */
			/* If the file does not exist, the connection string will be empty, the user will need to     */
			/* click the ... to the right of the connection string to generate a new string.  It will     */
			/* create the evi_ado.txt file. There is a Delete button that will remove this temporart .txt */
			/* file in the event the user wishes to remove that trace file.                               */
			try{
				readInd = 0;
				System::IO::StreamReader ^ sr = gcnew System::IO::StreamReader(fileName);
				sr->Peek();
				this->providerText->Text = sr->ReadLine();
				sr->Close();
			}
			catch (System::IO::FileNotFoundException ^ e)
			{
				/* Add possible log printing info, this can be used for future release.                        */
				/*MessageBox::Show(e->Message,"Invalid Input", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);*/
				String ^ i = e->Message;  /* Place holder to eliminate Warnings */
			}
			catch (System::IO::IOException ^ e)
			{
				/* Add possible log printing info, this can be used for future release.                        */
				/*MessageBox::Show(e->Message,"Invalid Input", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);*/
				String ^ i = e->Message;  /* Place holder to eliminate Warnings */
				MessageBox::Show(i,"Unhandled Exception", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
			}
			catch (...)
			{
				MessageBox::Show("Unhandled Exception","Unhandled Exception", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
			}
			__finally
			{
				readInd = 1;
			}
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~Form1()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  btnSubmit;
	private: System::Windows::Forms::Button^  btnClear;
	protected: 

	protected: 

	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::Label^  label2;
	private: System::Windows::Forms::TextBox^  providerText;
	private: System::Windows::Forms::TextBox^  queryText;


	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;

		/* List of variables that will be used within this class.  Necessary to be global so they can */
		/* be used throughout the application.  Filename currently works with Win7 operating systems  */
		/* but will be updated as soon as I can get my hands on the file system of a WinXP box.       */
		static System::String^ DSN = gcnew String("");
		static System::String^ query = gcnew String("");
		static System::String^ userProfile = Environment::GetEnvironmentVariable("USERPROFILE");
		static System::String^ fileName = userProfile + "\\AppData\\Local\\Temp\\evi_ado.txt";
		static System::String^ hiddenPass = gcnew String("");
		static System::String^ hiddenUser = gcnew String("");
		static System::String^ evi_rel = gcnew String("");
		static int readInd = 1;

	private: System::Windows::Forms::PictureBox^  pictureBox1;
	private: System::Windows::Forms::Button^  btnDelete;
	private: System::Windows::Forms::Label^  label3;
	private: System::Windows::Forms::CheckBox^  ckbxPassword;
	private: System::Windows::Forms::Label^  lblVersion;


	private: System::Windows::Forms::DataGridView^  dataGridView1;
	private: System::Windows::Forms::Label^  lblRowCount;
	private: System::Windows::Forms::LinkLabel^  lnkEvisions;
	private: System::Windows::Forms::Button^  btnADO;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			System::ComponentModel::ComponentResourceManager^  resources = (gcnew System::ComponentModel::ComponentResourceManager(Form1::typeid));
			System::Windows::Forms::DataGridViewCellStyle^  dataGridViewCellStyle2 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
			this->btnSubmit = (gcnew System::Windows::Forms::Button());
			this->btnClear = (gcnew System::Windows::Forms::Button());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->providerText = (gcnew System::Windows::Forms::TextBox());
			this->queryText = (gcnew System::Windows::Forms::TextBox());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			this->btnADO = (gcnew System::Windows::Forms::Button());
			this->btnDelete = (gcnew System::Windows::Forms::Button());
			this->label3 = (gcnew System::Windows::Forms::Label());
			this->ckbxPassword = (gcnew System::Windows::Forms::CheckBox());
			this->lblVersion = (gcnew System::Windows::Forms::Label());
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->lblRowCount = (gcnew System::Windows::Forms::Label());
			this->lnkEvisions = (gcnew System::Windows::Forms::LinkLabel());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^  >(this->pictureBox1))->BeginInit();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^  >(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// btnSubmit
			// 
			this->btnSubmit->Location = System::Drawing::Point(405, 575);
			this->btnSubmit->Name = L"btnSubmit";
			this->btnSubmit->Size = System::Drawing::Size(75, 23);
			this->btnSubmit->TabIndex = 3;
			this->btnSubmit->Tag = L"Execute Query";
			this->btnSubmit->Text = L"Execute";
			this->btnSubmit->UseVisualStyleBackColor = true;
			this->btnSubmit->Click += gcnew System::EventHandler(this, &Form1::btnSubmit_Click);
			// 
			// btnClear
			// 
			this->btnClear->Location = System::Drawing::Point(486, 575);
			this->btnClear->Name = L"btnClear";
			this->btnClear->Size = System::Drawing::Size(75, 23);
			this->btnClear->TabIndex = 4;
			this->btnClear->Tag = L"Clears Form Text";
			this->btnClear->Text = L"Clear";
			this->btnClear->UseVisualStyleBackColor = true;
			this->btnClear->Click += gcnew System::EventHandler(this, &Form1::btnClear_Click);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(9, 73);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(91, 13);
			this->label1->TabIndex = 8;
			this->label1->Text = L"Connection String";
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(13, 112);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(35, 13);
			this->label2->TabIndex = 8;
			this->label2->Text = L"Query";
			// 
			// providerText
			// 
			this->providerText->Location = System::Drawing::Point(12, 89);
			this->providerText->Name = L"providerText";
			this->providerText->Size = System::Drawing::Size(607, 20);
			this->providerText->TabIndex = 0;
			this->providerText->TextChanged += gcnew System::EventHandler(this, &Form1::providerText_TextChanged);
			// 
			// queryText
			// 
			this->queryText->AcceptsReturn = true;
			this->queryText->AllowDrop = true;
			this->queryText->Location = System::Drawing::Point(12, 128);
			this->queryText->Multiline = true;
			this->queryText->Name = L"queryText";
			this->queryText->ScrollBars = System::Windows::Forms::ScrollBars::Vertical;
			this->queryText->Size = System::Drawing::Size(628, 90);
			this->queryText->TabIndex = 2;
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackColor = System::Drawing::Color::White;
			this->pictureBox1->Image = (cli::safe_cast<System::Drawing::Image^  >(resources->GetObject(L"pictureBox1.Image")));
			this->pictureBox1->Location = System::Drawing::Point(0, 0);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(654, 70);
			this->pictureBox1->TabIndex = 9;
			this->pictureBox1->TabStop = false;
			// 
			// btnADO
			// 
			this->btnADO->Location = System::Drawing::Point(616, 89);
			this->btnADO->Name = L"btnADO";
			this->btnADO->Size = System::Drawing::Size(24, 21);
			this->btnADO->TabIndex = 1;
			this->btnADO->Text = L"...";
			this->btnADO->UseVisualStyleBackColor = true;
			this->btnADO->Click += gcnew System::EventHandler(this, &Form1::btnADO_Click);
			// 
			// btnDelete
			// 
			this->btnDelete->Location = System::Drawing::Point(567, 575);
			this->btnDelete->Name = L"btnDelete";
			this->btnDelete->Size = System::Drawing::Size(75, 23);
			this->btnDelete->TabIndex = 5;
			this->btnDelete->Tag = L"Deletes Resources";
			this->btnDelete->Text = L"Clear All";
			this->btnDelete->UseVisualStyleBackColor = true;
			this->btnDelete->Click += gcnew System::EventHandler(this, &Form1::btnDelete_Click);
			// 
			// label3
			// 
			this->label3->AutoSize = true;
			this->label3->Location = System::Drawing::Point(507, 73);
			this->label3->Name = L"label3";
			this->label3->Size = System::Drawing::Size(112, 13);
			this->label3->TabIndex = 8;
			this->label3->Text = L"Use Hidden Password";
			this->label3->TextAlign = System::Drawing::ContentAlignment::MiddleRight;
			// 
			// ckbxPassword
			// 
			this->ckbxPassword->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom) 
				| System::Windows::Forms::AnchorStyles::Left) 
				| System::Windows::Forms::AnchorStyles::Right));
			this->ckbxPassword->AutoSize = true;
			this->ckbxPassword->Location = System::Drawing::Point(625, 73);
			this->ckbxPassword->Name = L"ckbxPassword";
			this->ckbxPassword->Size = System::Drawing::Size(15, 14);
			this->ckbxPassword->TabIndex = 7;
			this->ckbxPassword->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
			this->ckbxPassword->UseVisualStyleBackColor = true;
			this->ckbxPassword->CheckedChanged += gcnew System::EventHandler(this, &Form1::ckbxPassword_CheckedChanged);
			// 
			// lblVersion
			// 
			this->lblVersion->AutoSize = true;
			this->lblVersion->Location = System::Drawing::Point(9, 580);
			this->lblVersion->Name = L"lblVersion";
			this->lblVersion->Size = System::Drawing::Size(76, 13);
			this->lblVersion->TabIndex = 8;
			this->lblVersion->Text = L"Release 1.24v";
			this->lblVersion->TextAlign = System::Drawing::ContentAlignment::MiddleLeft;
			// 
			// dataGridView1
			// 
			dataGridViewCellStyle2->Alignment = System::Windows::Forms::DataGridViewContentAlignment::TopLeft;
			dataGridViewCellStyle2->WrapMode = System::Windows::Forms::DataGridViewTriState::True;
			this->dataGridView1->AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
			this->dataGridView1->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::ColumnHeader;
			this->dataGridView1->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::AllHeaders;
			this->dataGridView1->CellBorderStyle = System::Windows::Forms::DataGridViewCellBorderStyle::None;
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Location = System::Drawing::Point(12, 239);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->ReadOnly = true;
			this->dataGridView1->Size = System::Drawing::Size(628, 330);
			this->dataGridView1->TabIndex = 8;
			// 
			// lblRowCount
			// 
			this->lblRowCount->AutoSize = true;
			this->lblRowCount->Location = System::Drawing::Point(13, 221);
			this->lblRowCount->Name = L"lblRowCount";
			this->lblRowCount->Size = System::Drawing::Size(0, 13);
			this->lblRowCount->TabIndex = 8;
			this->lblRowCount->TextAlign = System::Drawing::ContentAlignment::MiddleLeft;
			// 
			// lnkEvisions
			// 
			this->lnkEvisions->AutoSize = true;
			this->lnkEvisions->LinkColor = System::Drawing::Color::Black;
			this->lnkEvisions->Location = System::Drawing::Point(91, 580);
			this->lnkEvisions->Name = L"lnkEvisions";
			this->lnkEvisions->Size = System::Drawing::Size(88, 13);
			this->lnkEvisions->TabIndex = 6;
			this->lnkEvisions->TabStop = true;
			this->lnkEvisions->Text = L"Evisions Website";
			this->lnkEvisions->LinkClicked += gcnew System::Windows::Forms::LinkLabelLinkClickedEventHandler(this, &Form1::lnkEvisions_LinkClicked);
			// 
			// Form1
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(654, 602);
			this->Controls->Add(this->lnkEvisions);
			this->Controls->Add(this->lblRowCount);
			this->Controls->Add(this->lblVersion);
			this->Controls->Add(this->ckbxPassword);
			this->Controls->Add(this->label3);
			this->Controls->Add(this->dataGridView1);
			this->Controls->Add(this->btnDelete);
			this->Controls->Add(this->btnADO);
			this->Controls->Add(this->pictureBox1);
			this->Controls->Add(this->queryText);
			this->Controls->Add(this->providerText);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->btnClear);
			this->Controls->Add(this->btnSubmit);
			this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedSingle;
			this->MaximizeBox = false;
			this->MaximumSize = System::Drawing::Size(660, 630);
			this->MinimumSize = System::Drawing::Size(660, 630);
			this->Name = L"Form1";
			this->RightToLeft = System::Windows::Forms::RightToLeft::No;
			this->Text = L"ADO Tool";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^  >(this->pictureBox1))->EndInit();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^  >(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion


/* Submit button                                                                           */
/* This is the driving portion of the application.  The required elements include an ADO   */
/* connection string with/without a User ID/Password, a SQL statement in the form of a     */
/* query.  The connection string and the query are validated and, if either syntax is      */
/* incorrect, an alert will appear to the user with the error message.                     */
/*                                                                                         */
/* This also retrieves the data once the connection string and query have been validated.  */
/* The data retrieved will be Read Access Only.  The data cannot be modified, but each     */
/* column can be sorted in Ascending or Descending order.                                  */
private: System::Void btnSubmit_Click(System::Object^  sender, System::EventArgs^  e) {
			InitializeDataGridView();

			query = this->queryText->Text;
			DSN = this->providerText->Text;

			/* If the Connection String has no Password or User ID, use the credentials    */
			/* in the Hidden fields                                                        */
			if(ckbxPassword->Checked){
				if(!DSN->Contains("Password="))	{
					DSN += hiddenPass;
					if(!DSN->Contains("User ID="))	DSN += hiddenUser;
				}
			}

			OleDbConnection ^ myconn;

			try{
				/* Use the connection string and query to retrieve data from the datasource. */
				myconn = gcnew OleDbConnection(DSN);
				OleDbCommand ^ myCommand = gcnew OleDbCommand(query,myconn);
				myconn->Open();
				OleDbDataReader ^ myReader = myCommand->ExecuteReader();
				try 
				{
					/* Setup of the header fields, number of columns, number of rows, and row arrays */
					int row_count = 0;
					int printHeader = true;
					int col_count = myReader->FieldCount;
					dataGridView1->ColumnCount = col_count;
					dataGridView1->ColumnHeadersVisible = true;
					array<String^>^ rdata = gcnew array<String^>(col_count);
					array<Object^>^ rinsert = {rdata};

					/* Setup how the column headers will be formatted and displayed. */
					DataGridViewCellStyle ^ columnHeaderStyle = gcnew DataGridViewCellStyle;
					columnHeaderStyle->BackColor = Color::Aqua;
					columnHeaderStyle->Font = gcnew System::Drawing::Font( "Courier New",9,FontStyle::Bold );
					dataGridView1->ColumnHeadersDefaultCellStyle = columnHeaderStyle;

					/* Begin gathering Header/row information and display it on the Data Grid Viewer */
					while (myReader->Read()) {
						if(printHeader){
							for (int i = 0; col_count > i; i++){
								dataGridView1->Columns[ i ]->Name = myReader->GetName(i);
							}
							printHeader = false;
						}

						for (int i = 0; col_count > i; i++){
							rdata[i] = myReader->GetValue(i)->ToString();
						}

						System::Collections::IEnumerator^ myEnum = rinsert->GetEnumerator();
						myEnum->MoveNext();
						array<String^>^rowArray = safe_cast<array<String^>^>(myEnum->Current);
						dataGridView1->Rows->Add( rowArray );

						for (int i = 0; col_count > i; i++){
							rdata[i] = "";
						}

						row_count++;
					}
					this->lblRowCount->Text = "Row Count: " + row_count;

				}
				catch(ArgumentNullException ^e)
				{
					MessageBox::Show(e->Message,"Null Exception", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				}
				catch(ArgumentOutOfRangeException ^ e)
				{
					MessageBox::Show(e->Message,"Out of Range", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				}
				catch(Exception ^ e)
				{
					MessageBox::Show(e->Message,"Exception", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				}
				catch(...)
				{
					MessageBox::Show("Unhandled Exception","Unhandled Exception", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				}
				__finally 
				{
					myReader->Close();
				}

			}
			/* In the likely event, if the ADO connection string or query have invalid inputs, */
			/* The exceptions will be caught and displayed to the user.  The errors should be  */
			/* self-explanatory about what the issue is.                                       */
			catch (OleDbException ^ e)
			{
				if(e->Errors->default[0]->NativeError == 0)
					if(query == "")									// Empty query
						MessageBox::Show("Query is Empty.  Please enter a Query.","Invalid Query", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
					else											// No Value given for one or more required parameters
						MessageBox::Show(e->Message,"Parameter Missing", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				else												// Issues with portions of query
					MessageBox::Show(e->Message,"Invalid Query", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
			}
			catch(InvalidOperationException ^ e)
			{
				MessageBox::Show(e->Message,"Invalid Operation", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
			}
			catch(Exception ^ e)
			{
				MessageBox::Show(e->Message,"Exception", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
			}
			catch(...)
			{
				MessageBox::Show("Unhandled Exception","Unhandled Exception", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
			}
			__finally
			{
				try{
					myconn->Close();
				}
				catch(NullReferenceException ^ e)
				{
					MessageBox::Show(e->Message,"Invalid Input", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				}
			}
		 }


/* Clear button                                                                            */
/* This button will delete the temporary .txt document stored in the AppData\local\Temp    */
/* directory.  The user would use this in the event they wish to have an empty connection  */
/* string the next time it is launched. NOTE: The temporary .txt document is different for */
/* each individual user. This means that 2 users can have different default connection     */
/* depending on which user is logged into the machine.                                     */
private: System::Void btnClear_Click(System::Object^  sender, System::EventArgs^  e) {
			this->lblRowCount->Text = "";
			InitializeDataGridView();
		 }


/* Delete Button                                                                           */
/* This is the portion of the application that will clear all data on the form.  The       */
/* values that are cleared out include the connection string, the query, the Data Grid     */
/* Viewer, the hidden user and password, and the check box.                                */
private: System::Void btnDelete_Click(System::Object^  sender, System::EventArgs^  e) {
			this->ckbxPassword->Checked = false;
			this->providerText->Text = "";
			this->lblRowCount->Text = "";
			this->queryText->Text = "";
			hiddenPass = "";
			hiddenUser = "";
			query = "";
			DSN = "";
			InitializeDataGridView();
			System::IO::File::Delete(fileName);
		 }


/* InitializeDataGridView                                                                  */
/* This function will clear all of the columns and rows in the Grid.  It is called both    */
/* when the Clear buttons are used and every time the query executes.                      */
private:  System::Void InitializeDataGridView(){
			while(dataGridView1->ColumnCount > 0){
				dataGridView1->Columns->RemoveAt(0);
			}
		 }


/* ADO Connection String Loader                                                            */
/* This will load the Data Link properties window.  This allows the user to select their   */
/* connection provider, a datasource, a username and password.  Please NOTE: password is   */
/* REQUIRED to be saved in the datalink properties.  User can also skip adding username    */
/* and password by using the Hidden Password checkbox.                                     */
private: System::Void btnADO_Click(System::Object^  sender, System::EventArgs^  e) {
			 oledb::IDataSourceLocatorPtr	p_IDSL = NULL;
			 ADODB::_ConnectionPtr			p_conn = NULL;
			 String ^ DSN_provider = gcnew String("");
			 String ^ DSN_to_rest = gcnew String("");
			 int idx = 0;

			 p_IDSL.CreateInstance( __uuidof( oledb::DataLinks ));
			 p_conn.CreateInstance( "ADODB.Connection" );

			 IDispatch * p_Dispatch = NULL;
			 p_conn.QueryInterface( IID_IDispatch, (LPVOID*) & p_Dispatch );

			 if(p_IDSL->PromptEdit( &p_Dispatch )){
				 BSTR bstr = p_conn->ConnectionString;
				 DSN = gcnew String(bstr);
				 /* For use with Password-Protected Access database files */
				 if(DSN->Contains(";Password=") && (DSN->Contains(".mdb") || DSN->Contains(".accdb"))) {
					 idx = DSN->IndexOf("Password=");
					 DSN_provider = DSN->Substring(0, idx);
					 DSN_to_rest = DSN->Substring(idx);
					 DSN = DSN_provider + "Jet OLEDB:Database " + DSN_to_rest;
				 }
				 this->providerText->Text = DSN;
				 p_Dispatch->Release();
			 }
		 }


/* Hidden Password Checkbox                                                                */
/* This checkbox is used in the event that the user of the application wishes to store the */
/* user ID and password.  It is temporarily stored in the application when the checkbox is */
/* enabled.  if the clear button is clicked or the checkbox is unchecked, the values for   */
/* hidden username and password become empty strings.                                      */
/*                                                                                         */
/* If a password is not entered into the form then they the checkbox will become unchecked */
/* and nothing is stored.  Please NOTE: the password and user ID DO NOT get saved to the   */
/* temporary file that the DSN does, UNLESS it is included in the datalink properties tab. */
/* If any data is stored and the user wishes to remove it from their filesystem, they can  */
/* click the delete button.                                                                */
private: System::Void ckbxPassword_CheckedChanged(System::Object^  sender, System::EventArgs^  e) {
			 if(ckbxPassword->Checked == false){
				 hiddenPass = "";
				 hiddenUser = "";
				 return;
			 }

			 if(this->providerText->Text != ""){
				 if(ckbxPassword->Checked){
					DSN = this->providerText->Text;
					frmPassword ^ myForm2 = gcnew frmPassword();
					System::Windows::Forms::DialogResult rst = myForm2->ShowDialog(this);
					if(myForm2->txtPassword->Text->ToString() != "" ){
						/* For use with Password-Protected Access database files */
						if(DSN->Contains(".mdb") || DSN->Contains(".accdb")) {
							hiddenPass += ";Jet OLEDB:Database Password=" + myForm2->txtPassword->Text->ToString();
						}else{
							hiddenPass += ";Password=" + myForm2->txtPassword->Text->ToString();
						}

						/* Empty username appears to use default "Admin" user */
						if(myForm2->txtUserID->Text->ToString() != "")
							hiddenUser += ";User ID=" + myForm2->txtUserID->Text->ToString();

					}else{
						hiddenPass = "";
						hiddenUser = "";
						this->ckbxPassword->Checked = false;
						MessageBox::Show("Please specify a non-empty Username/Password","Invalid Input", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
					}
				}else{
					 hiddenPass = "";
					 hiddenUser = "";
				}
			}else{
				MessageBox::Show("Please specify an Connection String first","Invalid Input", MessageBoxButtons::OK,MessageBoxIcon::Asterisk);
				ckbxPassword->Checked = false;
			}
		 }



private: System::Void lnkEvisions_LinkClicked(System::Object^  sender, System::Windows::Forms::LinkLabelLinkClickedEventArgs^  e) {
			String ^ target= "http://www.evisions.com";

			try{
				System::Diagnostics::Process::Start(target);
			}
			catch (System::ComponentModel::Win32Exception ^ noBrowser) 
			{
				if (noBrowser->ErrorCode==-2147467259)
				MessageBox::Show(noBrowser->Message);
			}
			catch (System::Exception ^ other)
			{
				MessageBox::Show(other->Message);
			}
		 }



private: System::Void providerText_TextChanged(System::Object^  sender, System::EventArgs^  e) {
			if(readInd == 1){
				DSN = this->providerText->Text;
				printToFile();
			}
		 }



private: System::Void printToFile(){
			System::IO::StreamWriter ^ sw = gcnew System::IO::StreamWriter(fileName);
			sw->WriteLine(DSN);
			sw->Close();
		 }
};
}




