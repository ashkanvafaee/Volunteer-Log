/****************************************************************************************
New Volunteer Registration Form
*****************************************************************************************/
#pragma once
#include <string>
#include "Date.h"





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
	/// Summary for DataBox
	/// </summary>
	public ref class DataBox : public System::Windows::Forms::Form
	{
	public:
		DataBox(void)
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
		~DataBox()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::TextBox^  name;
	private: System::Windows::Forms::TextBox^  email;



	protected:


	private: System::Windows::Forms::Label^  nameLabel;
	private: System::Windows::Forms::Label^  phoneLabel;






	private: System::Windows::Forms::TextBox^  phone;


	private: System::Windows::Forms::Label^  label2;
	private: System::Windows::Forms::Button^  button2;
	private: System::Windows::Forms::PictureBox^  pictureBox1;

	protected:

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
			System::ComponentModel::ComponentResourceManager^  resources = (gcnew System::ComponentModel::ComponentResourceManager(DataBox::typeid));
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->name = (gcnew System::Windows::Forms::TextBox());
			this->email = (gcnew System::Windows::Forms::TextBox());
			this->nameLabel = (gcnew System::Windows::Forms::Label());
			this->phoneLabel = (gcnew System::Windows::Forms::Label());
			this->phone = (gcnew System::Windows::Forms::TextBox());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->SuspendLayout();
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label1->Location = System::Drawing::Point(136, 28);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(331, 33);
			this->label1->TabIndex = 0;
			this->label1->Text = L"Please Enter Information";
			// 
			// name
			// 
			this->name->Location = System::Drawing::Point(260, 101);
			this->name->Name = L"name";
			this->name->Size = System::Drawing::Size(236, 20);
			this->name->TabIndex = 1;
			this->name->Enter += gcnew System::EventHandler(this, &DataBox::name_Enter);
			this->name->Leave += gcnew System::EventHandler(this, &DataBox::name_Leave);
			// 
			// email
			// 
			this->email->Location = System::Drawing::Point(260, 154);
			this->email->Name = L"email";
			this->email->Size = System::Drawing::Size(236, 20);
			this->email->TabIndex = 2;
			this->email->Enter += gcnew System::EventHandler(this, &DataBox::email_Enter);
			this->email->Leave += gcnew System::EventHandler(this, &DataBox::email_Leave);
			// 
			// nameLabel
			// 
			this->nameLabel->AutoSize = true;
			this->nameLabel->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->nameLabel->Location = System::Drawing::Point(25, 87);
			this->nameLabel->Name = L"nameLabel";
			this->nameLabel->Size = System::Drawing::Size(104, 33);
			this->nameLabel->TabIndex = 3;
			this->nameLabel->Text = L"Name:";
			// 
			// phoneLabel
			// 
			this->phoneLabel->AutoSize = true;
			this->phoneLabel->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->phoneLabel->Location = System::Drawing::Point(25, 198);
			this->phoneLabel->Name = L"phoneLabel";
			this->phoneLabel->Size = System::Drawing::Size(219, 33);
			this->phoneLabel->TabIndex = 4;
			this->phoneLabel->Text = L"Phone Number:";
			// 
			// phone
			// 
			this->phone->Location = System::Drawing::Point(260, 211);
			this->phone->Name = L"phone";
			this->phone->Size = System::Drawing::Size(236, 20);
			this->phone->TabIndex = 6;
			this->phone->Enter += gcnew System::EventHandler(this, &DataBox::phone_Enter);
			this->phone->Leave += gcnew System::EventHandler(this, &DataBox::phone_Leave);
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label2->Location = System::Drawing::Point(25, 141);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(92, 33);
			this->label2->TabIndex = 7;
			this->label2->Text = L"Email:";
			// 
			// button2
			// 
			this->button2->BackColor = System::Drawing::Color::Transparent;
			this->button2->Location = System::Drawing::Point(260, 277);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(75, 23);
			this->button2->TabIndex = 8;
			this->button2->Text = L"Done";
			this->button2->UseVisualStyleBackColor = false;
			this->button2->Click += gcnew System::EventHandler(this, &DataBox::button2_Click);
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackColor = System::Drawing::Color::Transparent;
			this->pictureBox1->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.Image")));
			this->pictureBox1->Location = System::Drawing::Point(502, 246);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(113, 88);
			this->pictureBox1->SizeMode = System::Windows::Forms::PictureBoxSizeMode::StretchImage;
			this->pictureBox1->TabIndex = 11;
			this->pictureBox1->TabStop = false;
			// 
			// DataBox
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::Orange;
			this->ClientSize = System::Drawing::Size(609, 336);
			this->Controls->Add(this->pictureBox1);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->phone);
			this->Controls->Add(this->email);
			this->Controls->Add(this->name);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->phoneLabel);
			this->Controls->Add(this->nameLabel);
			this->Controls->Add(this->label1);
			this->Name = L"DataBox";
			this->Text = L"DataBox";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion


private: System::Void button2_Click(System::Object^  sender, System::EventArgs^  e) {

	if (name->Text == "") {
		MessageBox::Show(Owner, "Please Enter Name", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);
	}
	else {



		NetOffice::ExcelApi::Application^ oApp;
		NetOffice::ExcelApi::Worksheet^ oSheet;
		NetOffice::ExcelApi::Workbook^ oBook;

		oApp = gcnew NetOffice::ExcelApi::Application();
		oBook = oApp->Workbooks->Open("C:\\Volunteer Log");
		oSheet = (NetOffice::ExcelApi::Worksheet^)oBook->Worksheets[1];

		//if (oSheet->Cells[1,8]->Value == nullptr) {																	//determines if cell is empty
		//	int x=3;
		//}


		System::String^ str = oSheet->Cells[1, 8]->Value->ToString();
		int i = int::Parse(str);
		i = i * 6;



		oSheet->Cells[2 + i, 4]->Value = "Date";
		oSheet->Cells[3 + i, 4]->Value = "Time In";
		oSheet->Cells[4 + i, 4]->Value = "Time Out";
		oSheet->Cells[5 + i, 4]->Value = "Hours";
		oSheet->Cells[5 + i, 1]->Value = "Total:";



		for (int k = 0; k < 4; k++) {
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlInsideVertical]->LineStyle = XlLineStyle::xlContinuous;
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlInsideVertical]->Weight = 2;
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlEdgeBottom]->LineStyle = XlLineStyle::xlContinuous;
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlEdgeBottom]->Weight = 2;
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlEdgeTop]->LineStyle = XlLineStyle::xlContinuous;
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlEdgeTop]->Weight = 2;
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlEdgeLeft]->LineStyle = XlLineStyle::xlContinuous;
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlEdgeLeft]->Weight = 2;
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlEdgeRight]->LineStyle = XlLineStyle::xlContinuous;
			oSheet->Rows[2 + i + k]->Borders[XlBordersIndex::xlEdgeRight]->Weight = 2;
		}





		oSheet->Range(oSheet->Cells[2 + i, 1], oSheet->Cells[2 + i, 3])->Merge(Type::Missing);						//Merges
		oSheet->Range(oSheet->Cells[3 + i, 1], oSheet->Cells[3 + i, 3])->Merge(Type::Missing);
		oSheet->Range(oSheet->Cells[4 + i, 1], oSheet->Cells[4 + i, 3])->Merge(Type::Missing);
		oSheet->Range(oSheet->Cells[5 + i, 1], oSheet->Cells[5 + i, 3])->Merge(Type::Missing);


		oSheet->Cells[3 + i, 1]->HorizontalAlignment = XlHAlign::xlHAlignLeft;										//Aligns phone number to the left


		oSheet->Cells[2 + i, 1]->Value = name->Text;
		oSheet->Cells[3 + i, 1]->Value = phone->Text;
		oSheet->Cells[4 + i, 1]->Value = email->Text;

		int namePosition = i;																						//Sets name position

		i = i + 6;
		oSheet->Cells[1, 8]->Value = i / 6;








		oBook->Save();
		oApp->Quit();
		oApp->DisposeChildInstances();
		this->Close();



		Date^ b = gcnew Date;
		b->Show();
		b->namePosition = namePosition;
		delete oApp;
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
