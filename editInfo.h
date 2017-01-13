#pragma once

namespace Data_Base_2 {

	using namespace NetOffice::ExcelApi;
	using namespace NetOffice::ExcelApi::Enums;
	using namespace NetOffice::ExcelApi::Tools::Utils;
	using namespace System::IO;
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Summary for editInfo
	/// </summary>
	public ref class editInfo : public System::Windows::Forms::Form
	{
	public:
		editInfo(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}
	private: System::Windows::Forms::PictureBox^  pictureBox1;
	public:
		static int nameIndex = 0;

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~editInfo()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  button2;
	public:
	 System::Windows::Forms::TextBox^  phone;
	 System::Windows::Forms::TextBox^  email;
	 System::Windows::Forms::TextBox^  name;
	 System::Windows::Forms::Label^  label2;
	 System::Windows::Forms::Label^  phoneLabel;
	 System::Windows::Forms::Label^  nameLabel;
	 System::Windows::Forms::Label^  label1;

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
			System::ComponentModel::ComponentResourceManager^  resources = (gcnew System::ComponentModel::ComponentResourceManager(editInfo::typeid));
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->phone = (gcnew System::Windows::Forms::TextBox());
			this->email = (gcnew System::Windows::Forms::TextBox());
			this->name = (gcnew System::Windows::Forms::TextBox());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->phoneLabel = (gcnew System::Windows::Forms::Label());
			this->nameLabel = (gcnew System::Windows::Forms::Label());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->SuspendLayout();
			// 
			// button2
			// 
			this->button2->BackColor = System::Drawing::Color::Transparent;
			this->button2->Location = System::Drawing::Point(237, 298);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(130, 27);
			this->button2->TabIndex = 16;
			this->button2->Text = L"Save Information";
			this->button2->UseVisualStyleBackColor = false;
			this->button2->Click += gcnew System::EventHandler(this, &editInfo::button2_Click);
			// 
			// phone
			// 
			this->phone->Location = System::Drawing::Point(297, 230);
			this->phone->Name = L"phone";
			this->phone->Size = System::Drawing::Size(236, 20);
			this->phone->TabIndex = 14;
			this->phone->Enter += gcnew System::EventHandler(this, &editInfo::phone_Enter);
			this->phone->Leave += gcnew System::EventHandler(this, &editInfo::phone_Leave);
			// 
			// email
			// 
			this->email->Location = System::Drawing::Point(297, 173);
			this->email->Name = L"email";
			this->email->Size = System::Drawing::Size(236, 20);
			this->email->TabIndex = 11;
			this->email->Enter += gcnew System::EventHandler(this, &editInfo::email_Enter);
			this->email->Leave += gcnew System::EventHandler(this, &editInfo::email_Leave);
			// 
			// name
			// 
			this->name->Location = System::Drawing::Point(297, 120);
			this->name->Name = L"name";
			this->name->Size = System::Drawing::Size(236, 20);
			this->name->TabIndex = 10;
			this->name->Enter += gcnew System::EventHandler(this, &editInfo::name_Enter);
			this->name->Leave += gcnew System::EventHandler(this, &editInfo::name_Leave);
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label2->Location = System::Drawing::Point(62, 160);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(92, 33);
			this->label2->TabIndex = 15;
			this->label2->Text = L"Email:";
			// 
			// phoneLabel
			// 
			this->phoneLabel->AutoSize = true;
			this->phoneLabel->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->phoneLabel->Location = System::Drawing::Point(62, 217);
			this->phoneLabel->Name = L"phoneLabel";
			this->phoneLabel->Size = System::Drawing::Size(219, 33);
			this->phoneLabel->TabIndex = 13;
			this->phoneLabel->Text = L"Phone Number:";
			// 
			// nameLabel
			// 
			this->nameLabel->AutoSize = true;
			this->nameLabel->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->nameLabel->Location = System::Drawing::Point(62, 106);
			this->nameLabel->Name = L"nameLabel";
			this->nameLabel->Size = System::Drawing::Size(104, 33);
			this->nameLabel->TabIndex = 12;
			this->nameLabel->Text = L"Name:";
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label1->Location = System::Drawing::Point(173, 47);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(285, 33);
			this->label1->TabIndex = 9;
			this->label1->Text = L"Member Information";
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackColor = System::Drawing::Color::Transparent;
			this->pictureBox1->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.Image")));
			this->pictureBox1->Location = System::Drawing::Point(506, 265);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(113, 88);
			this->pictureBox1->SizeMode = System::Windows::Forms::PictureBoxSizeMode::StretchImage;
			this->pictureBox1->TabIndex = 17;
			this->pictureBox1->TabStop = false;
			// 
			// editInfo
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::Orange;
			this->ClientSize = System::Drawing::Size(619, 353);
			this->Controls->Add(this->pictureBox1);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->phone);
			this->Controls->Add(this->email);
			this->Controls->Add(this->name);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->phoneLabel);
			this->Controls->Add(this->nameLabel);
			this->Controls->Add(this->label1);
			this->Name = L"editInfo";
			this->Text = L"editInfo";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion

	private: System::Void button2_Click(System::Object^  sender, System::EventArgs^  e) {
		int i = nameIndex;

		if (name->Text == "") {
			MessageBox::Show(Owner, "Please enter a name.", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);
		}
		else {



			NetOffice::ExcelApi::Application^ oApp;
			NetOffice::ExcelApi::Worksheet^ oSheet;
			NetOffice::ExcelApi::Workbook^ oBook;

			oApp = gcnew NetOffice::ExcelApi::Application();
			oBook = oApp->Workbooks->Open("C:\\Volunteer Log");
			oSheet = (NetOffice::ExcelApi::Worksheet^) oBook->ActiveSheet;


			oSheet->Cells[i + 2, 1]->Value = name->Text;

			oSheet->Cells[i + 4, 1]->Value = email->Text;

			oSheet->Cells[i + 3, 1]->Value = phone->Text;


			oSheet->Cells[i + 3 , 1]->HorizontalAlignment = XlHAlign::xlHAlignLeft;												//Aligns phone number to the left



			oBook->Save();
			oApp->Quit();
			oApp->DisposeChildInstances();
			delete oApp;
			this->Close();

			MessageBox::Show(Owner, "Information Saved!", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);


		}


	}









private: System::Void name_Enter(System::Object^  sender, System::EventArgs^  e) {
	name->BackColor = System::Drawing::Color::LightGreen;
}
private: System::Void name_Leave(System::Object^  sender, System::EventArgs^  e) {
	name->BackColor = System::Drawing::Color::White;
}
private: System::Void email_Enter(System::Object^  sender, System::EventArgs^  e) {
	email->BackColor = System::Drawing::Color::LightGreen;
}
private: System::Void email_Leave(System::Object^  sender, System::EventArgs^  e) {
	email->BackColor = System::Drawing::Color::White;
}
private: System::Void phone_Enter(System::Object^  sender, System::EventArgs^  e) {
	phone->BackColor = System::Drawing::Color::LightGreen;
}
private: System::Void phone_Leave(System::Object^  sender, System::EventArgs^  e) {
	phone->BackColor = System::Drawing::Color::White;
}




};
}
