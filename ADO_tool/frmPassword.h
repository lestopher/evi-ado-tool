#pragma once

namespace ADO_tool {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Summary for frmPassword
	/// </summary>
	public ref class frmPassword : public System::Windows::Forms::Form
	{
	public:
		frmPassword(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~frmPassword()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  btnSubmit;
	protected: 
	private: System::Windows::Forms::Label^  label1;
	public: System::Windows::Forms::TextBox^  txtPassword;
	public: System::Windows::Forms::TextBox^  txtUserID;
	private: System::Windows::Forms::Label^  label2;
	public: 



	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			this->btnSubmit = (gcnew System::Windows::Forms::Button());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->txtPassword = (gcnew System::Windows::Forms::TextBox());
			this->txtUserID = (gcnew System::Windows::Forms::TextBox());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->SuspendLayout();
			// 
			// btnSubmit
			// 
			this->btnSubmit->Location = System::Drawing::Point(169, 90);
			this->btnSubmit->Name = L"btnSubmit";
			this->btnSubmit->Size = System::Drawing::Size(75, 23);
			this->btnSubmit->TabIndex = 2;
			this->btnSubmit->Text = L"Submit";
			this->btnSubmit->UseVisualStyleBackColor = true;
			this->btnSubmit->Click += gcnew System::EventHandler(this, &frmPassword::btnSubmit_Click);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(9, 9);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(145, 13);
			this->label1->TabIndex = 1;
			this->label1->Text = L"User ID for connection string:";
			// 
			// txtPassword
			// 
			this->txtPassword->Location = System::Drawing::Point(12, 64);
			this->txtPassword->Name = L"txtPassword";
			this->txtPassword->Size = System::Drawing::Size(232, 20);
			this->txtPassword->TabIndex = 1;
			this->txtPassword->UseSystemPasswordChar = true;
			// 
			// txtUserID
			// 
			this->txtUserID->Location = System::Drawing::Point(12, 25);
			this->txtUserID->Name = L"txtUserID";
			this->txtUserID->Size = System::Drawing::Size(232, 20);
			this->txtUserID->TabIndex = 0;
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(12, 48);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(155, 13);
			this->label2->TabIndex = 3;
			this->label2->Text = L"Password for connection string:";
			// 
			// frmPassword
			// 
			this->AllowDrop = true;
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(259, 122);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->txtUserID);
			this->Controls->Add(this->txtPassword);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->btnSubmit);
			this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedSingle;
			this->MaximizeBox = false;
			this->MaximumSize = System::Drawing::Size(275, 160);
			this->MinimumSize = System::Drawing::Size(275, 160);
			this->Name = L"frmPassword";
			this->Text = L"ADO Credentials";
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void btnSubmit_Click(System::Object^  sender, System::EventArgs^  e) {
				 this->Hide();
			 }
	};
}
