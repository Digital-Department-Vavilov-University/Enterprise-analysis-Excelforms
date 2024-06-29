#pragma once
#include <iostream>

using namespace std;

namespace Project14 {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace Microsoft::Office::Interop;

	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^ buttonLoad;
	protected:
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::OpenFileDialog^ openFileDialog1;
	private: System::Windows::Forms::SaveFileDialog^ saveFileDialog1;
	private: System::Windows::Forms::Button^ buttonSave;
	private: System::Windows::Forms::Button^ buttonProfitability;



	private: System::Windows::Forms::Label^ labelProfitability;
	private: System::Windows::Forms::Button^ buttonProfitRate;
	private: System::Windows::Forms::Label^ labelProfitRate;
	private: System::Windows::Forms::Label^ labelCapitalLaborRatio;
	private: System::Windows::Forms::Button^ buttonCapitalLaborRatio;
	private: System::Windows::Forms::Label^ labelKo;
	private: System::Windows::Forms::Button^ buttonKo;
	private: System::Windows::Forms::Label^ label1;
	private: System::Windows::Forms::Label^ label2;
	private: System::Windows::Forms::Label^ label3;
	private: System::Windows::Forms::Label^ label4;




	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle1 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
			System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle2 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
			System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle3 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
			System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle4 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MyForm::typeid));
			this->buttonLoad = (gcnew System::Windows::Forms::Button());
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->openFileDialog1 = (gcnew System::Windows::Forms::OpenFileDialog());
			this->saveFileDialog1 = (gcnew System::Windows::Forms::SaveFileDialog());
			this->buttonSave = (gcnew System::Windows::Forms::Button());
			this->buttonProfitability = (gcnew System::Windows::Forms::Button());
			this->labelProfitability = (gcnew System::Windows::Forms::Label());
			this->buttonProfitRate = (gcnew System::Windows::Forms::Button());
			this->labelProfitRate = (gcnew System::Windows::Forms::Label());
			this->labelCapitalLaborRatio = (gcnew System::Windows::Forms::Label());
			this->buttonCapitalLaborRatio = (gcnew System::Windows::Forms::Button());
			this->labelKo = (gcnew System::Windows::Forms::Label());
			this->buttonKo = (gcnew System::Windows::Forms::Button());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->label3 = (gcnew System::Windows::Forms::Label());
			this->label4 = (gcnew System::Windows::Forms::Label());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// buttonLoad
			// 
			this->buttonLoad->Location = System::Drawing::Point(911, 332);
			this->buttonLoad->Margin = System::Windows::Forms::Padding(2);
			this->buttonLoad->Name = L"buttonLoad";
			this->buttonLoad->Size = System::Drawing::Size(106, 46);
			this->buttonLoad->TabIndex = 0;
			this->buttonLoad->Text = L"Загрузить";
			this->buttonLoad->UseVisualStyleBackColor = true;
			this->buttonLoad->Click += gcnew System::EventHandler(this, &MyForm::buttonLoad_Click);
			// 
			// dataGridView1
			// 
			this->dataGridView1->BackgroundColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(224)),
				static_cast<System::Int32>(static_cast<System::Byte>(224)), static_cast<System::Int32>(static_cast<System::Byte>(224)));
			dataGridViewCellStyle1->Alignment = System::Windows::Forms::DataGridViewContentAlignment::MiddleLeft;
			dataGridViewCellStyle1->BackColor = System::Drawing::SystemColors::ControlDark;
			dataGridViewCellStyle1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Regular,
				System::Drawing::GraphicsUnit::Point, static_cast<System::Byte>(204)));
			dataGridViewCellStyle1->ForeColor = System::Drawing::SystemColors::WindowText;
			dataGridViewCellStyle1->SelectionBackColor = System::Drawing::SystemColors::Highlight;
			dataGridViewCellStyle1->SelectionForeColor = System::Drawing::SystemColors::Control;
			dataGridViewCellStyle1->WrapMode = System::Windows::Forms::DataGridViewTriState::True;
			this->dataGridView1->ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Cursor = System::Windows::Forms::Cursors::Default;
			dataGridViewCellStyle2->Alignment = System::Windows::Forms::DataGridViewContentAlignment::MiddleLeft;
			dataGridViewCellStyle2->BackColor = System::Drawing::SystemColors::Menu;
			dataGridViewCellStyle2->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			dataGridViewCellStyle2->ForeColor = System::Drawing::SystemColors::ControlText;
			dataGridViewCellStyle2->SelectionBackColor = System::Drawing::SystemColors::Highlight;
			dataGridViewCellStyle2->SelectionForeColor = System::Drawing::SystemColors::ControlLight;
			dataGridViewCellStyle2->WrapMode = System::Windows::Forms::DataGridViewTriState::False;
			this->dataGridView1->DefaultCellStyle = dataGridViewCellStyle2;
			this->dataGridView1->EnableHeadersVisualStyles = false;
			this->dataGridView1->GridColor = System::Drawing::SystemColors::Control;
			this->dataGridView1->Location = System::Drawing::Point(18, 29);
			this->dataGridView1->Margin = System::Windows::Forms::Padding(2);
			this->dataGridView1->Name = L"dataGridView1";
			dataGridViewCellStyle3->Alignment = System::Windows::Forms::DataGridViewContentAlignment::MiddleLeft;
			dataGridViewCellStyle3->BackColor = System::Drawing::SystemColors::Control;
			dataGridViewCellStyle3->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Regular,
				System::Drawing::GraphicsUnit::Point, static_cast<System::Byte>(204)));
			dataGridViewCellStyle3->ForeColor = System::Drawing::SystemColors::WindowText;
			dataGridViewCellStyle3->SelectionBackColor = System::Drawing::SystemColors::Highlight;
			dataGridViewCellStyle3->SelectionForeColor = System::Drawing::SystemColors::Control;
			dataGridViewCellStyle3->WrapMode = System::Windows::Forms::DataGridViewTriState::True;
			this->dataGridView1->RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
			this->dataGridView1->RowHeadersWidth = 51;
			dataGridViewCellStyle4->BackColor = System::Drawing::Color::White;
			dataGridViewCellStyle4->ForeColor = System::Drawing::Color::Black;
			dataGridViewCellStyle4->SelectionBackColor = System::Drawing::Color::Silver;
			dataGridViewCellStyle4->SelectionForeColor = System::Drawing::Color::Black;
			this->dataGridView1->RowsDefaultCellStyle = dataGridViewCellStyle4;
			this->dataGridView1->RowTemplate->Height = 24;
			this->dataGridView1->Size = System::Drawing::Size(999, 280);
			this->dataGridView1->TabIndex = 1;
			// 
			// openFileDialog1
			// 
			this->openFileDialog1->FileName = L"openFileDialog1";
			// 
			// buttonSave
			// 
			this->buttonSave->Location = System::Drawing::Point(911, 408);
			this->buttonSave->Margin = System::Windows::Forms::Padding(2);
			this->buttonSave->Name = L"buttonSave";
			this->buttonSave->Size = System::Drawing::Size(106, 46);
			this->buttonSave->TabIndex = 2;
			this->buttonSave->Text = L"Сохранить";
			this->buttonSave->UseVisualStyleBackColor = true;
			this->buttonSave->Click += gcnew System::EventHandler(this, &MyForm::buttonSave_Click);
			// 
			// buttonProfitability
			// 
			this->buttonProfitability->Location = System::Drawing::Point(42, 347);
			this->buttonProfitability->Margin = System::Windows::Forms::Padding(2);
			this->buttonProfitability->Name = L"buttonProfitability";
			this->buttonProfitability->Size = System::Drawing::Size(139, 31);
			this->buttonProfitability->TabIndex = 3;
			this->buttonProfitability->Text = L"Рентабельность";
			this->buttonProfitability->UseVisualStyleBackColor = true;
			this->buttonProfitability->Click += gcnew System::EventHandler(this, &MyForm::buttonRent_Click);
			// 
			// labelProfitability
			// 
			this->labelProfitability->AutoSize = true;
			this->labelProfitability->Location = System::Drawing::Point(196, 349);
			this->labelProfitability->Margin = System::Windows::Forms::Padding(2, 0, 2, 0);
			this->labelProfitability->Name = L"labelProfitability";
			this->labelProfitability->Size = System::Drawing::Size(0, 13);
			this->labelProfitability->TabIndex = 6;
			// 
			// buttonProfitRate
			// 
			this->buttonProfitRate->Location = System::Drawing::Point(42, 422);
			this->buttonProfitRate->Margin = System::Windows::Forms::Padding(2);
			this->buttonProfitRate->Name = L"buttonProfitRate";
			this->buttonProfitRate->Size = System::Drawing::Size(139, 31);
			this->buttonProfitRate->TabIndex = 7;
			this->buttonProfitRate->Text = L"Норма прибыли";
			this->buttonProfitRate->UseVisualStyleBackColor = true;
			this->buttonProfitRate->Click += gcnew System::EventHandler(this, &MyForm::buttonProfitRate_Click);
			// 
			// labelProfitRate
			// 
			this->labelProfitRate->AutoSize = true;
			this->labelProfitRate->Location = System::Drawing::Point(196, 431);
			this->labelProfitRate->Margin = System::Windows::Forms::Padding(2, 0, 2, 0);
			this->labelProfitRate->Name = L"labelProfitRate";
			this->labelProfitRate->Size = System::Drawing::Size(0, 13);
			this->labelProfitRate->TabIndex = 8;
			// 
			// labelCapitalLaborRatio
			// 
			this->labelCapitalLaborRatio->AutoSize = true;
			this->labelCapitalLaborRatio->Location = System::Drawing::Point(542, 349);
			this->labelCapitalLaborRatio->Margin = System::Windows::Forms::Padding(2, 0, 2, 0);
			this->labelCapitalLaborRatio->Name = L"labelCapitalLaborRatio";
			this->labelCapitalLaborRatio->Size = System::Drawing::Size(0, 13);
			this->labelCapitalLaborRatio->TabIndex = 10;
			// 
			// buttonCapitalLaborRatio
			// 
			this->buttonCapitalLaborRatio->Location = System::Drawing::Point(389, 349);
			this->buttonCapitalLaborRatio->Margin = System::Windows::Forms::Padding(2);
			this->buttonCapitalLaborRatio->Name = L"buttonCapitalLaborRatio";
			this->buttonCapitalLaborRatio->Size = System::Drawing::Size(139, 31);
			this->buttonCapitalLaborRatio->TabIndex = 9;
			this->buttonCapitalLaborRatio->Text = L"Фондовооруженность";
			this->buttonCapitalLaborRatio->UseVisualStyleBackColor = true;
			this->buttonCapitalLaborRatio->Click += gcnew System::EventHandler(this, &MyForm::buttonCapitalLaborRatio_Click);
			// 
			// labelKo
			// 
			this->labelKo->AutoSize = true;
			this->labelKo->Location = System::Drawing::Point(542, 422);
			this->labelKo->Margin = System::Windows::Forms::Padding(2, 0, 2, 0);
			this->labelKo->Name = L"labelKo";
			this->labelKo->Size = System::Drawing::Size(0, 13);
			this->labelKo->TabIndex = 12;
			// 
			// buttonKo
			// 
			this->buttonKo->Location = System::Drawing::Point(389, 415);
			this->buttonKo->Margin = System::Windows::Forms::Padding(2);
			this->buttonKo->Name = L"buttonKo";
			this->buttonKo->Size = System::Drawing::Size(139, 39);
			this->buttonKo->TabIndex = 11;
			this->buttonKo->Text = L"Коэфициент оборачиваемости";
			this->buttonKo->UseVisualStyleBackColor = true;
			this->buttonKo->Click += gcnew System::EventHandler(this, &MyForm::buttonKo_Click);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Italic, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label1->Location = System::Drawing::Point(51, 332);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(115, 13);
			this->label1->TabIndex = 13;
			this->label1->Text = L"Загрузите форму №2";
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Italic, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label2->Location = System::Drawing::Point(51, 407);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(115, 13);
			this->label2->TabIndex = 14;
			this->label2->Text = L"Загрузите форму №1";
			// 
			// label3
			// 
			this->label3->AutoSize = true;
			this->label3->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Italic, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label3->Location = System::Drawing::Point(401, 334);
			this->label3->Name = L"label3";
			this->label3->Size = System::Drawing::Size(115, 13);
			this->label3->TabIndex = 15;
			this->label3->Text = L"Загрузите форму №5";
			// 
			// label4
			// 
			this->label4->AutoSize = true;
			this->label4->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Italic, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label4->Location = System::Drawing::Point(401, 400);
			this->label4->Name = L"label4";
			this->label4->Size = System::Drawing::Size(115, 13);
			this->label4->TabIndex = 16;
			this->label4->Text = L"Загрузите форму №6";
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::Honeydew;
			this->BackgroundImageLayout = System::Windows::Forms::ImageLayout::None;
			this->ClientSize = System::Drawing::Size(1041, 489);
			this->Controls->Add(this->label4);
			this->Controls->Add(this->label3);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->labelKo);
			this->Controls->Add(this->buttonKo);
			this->Controls->Add(this->labelCapitalLaborRatio);
			this->Controls->Add(this->buttonCapitalLaborRatio);
			this->Controls->Add(this->labelProfitRate);
			this->Controls->Add(this->buttonProfitRate);
			this->Controls->Add(this->labelProfitability);
			this->Controls->Add(this->buttonProfitability);
			this->Controls->Add(this->buttonSave);
			this->Controls->Add(this->dataGridView1);
			this->Controls->Add(this->buttonLoad);
			this->Cursor = System::Windows::Forms::Cursors::Default;
			this->ForeColor = System::Drawing::SystemColors::ControlText;
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->Margin = System::Windows::Forms::Padding(2);
			this->Name = L"MyForm";
			this->Text = L"Анализ предприятия по формам";
			this->TransparencyKey = System::Drawing::Color::Transparent;
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
		String^ filePath;
		Excel::Application^ exlApp;
		Excel::Workbook^ exlBook;
		Excel::Worksheet^ exlSheet;
		Excel::Range^ exlRange;
		DataTable^ dt;

		bool OpenFile()
		{
			openFileDialog1->Filter = "(excel file(*.xlsx)| *.xlsx";
			openFileDialog1->RestoreDirectory = true;
			try
			{
				if (openFileDialog1->ShowDialog() == System::Windows::Forms::DialogResult::OK)
				{
					filePath = openFileDialog1->FileName;
					exlApp = gcnew Excel::ApplicationClass();
					exlBook = exlApp->Workbooks->Open(filePath, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
					exlSheet = safe_cast<Excel::Worksheet^>(exlBook->Sheets[1]);
					exlRange = exlSheet->UsedRange;
					dt = gcnew DataTable();
					return true;
				}
			}
			catch (Exception^ exception)
			{
				MessageBox::Show(this, "не удалось открыть файл", exception->Message,
					MessageBoxButtons::OK, MessageBoxIcon::Error);
				return false;
			}
		}
		bool CheckName(String^ nameTable, int x, int y)
		{
			if (OpenFile())
			{
				if ((safe_cast<Excel::Range^>(exlRange->Cells[x, y])->Value2)->ToString() == nameTable) return true;
				else
				{
					MessageBox::Show("Это не "+ nameTable, "Ошибка файла",
						MessageBoxButtons::OK, MessageBoxIcon::Error);
					return false;
				}
			}
			else { return false; }

		}
		String^ RoundDouble(double number, int capacity)
		{
			String^ numberText = number.ToString();
			String^ roundNumberText;
			for (int i = 0; i < numberText->Length; i++)
			{
				roundNumberText += numberText[i];
				if (numberText[i] == ',')
				{
					for (int j = 1; j < (capacity + 1); j++)
					{
						roundNumberText += numberText[i + j];
					}
					break;
				}
			}
				return roundNumberText;
		}

	private: System::Void buttonLoad_Click(System::Object^ sender, System::EventArgs^ e) {
		if (OpenFile()) {
			for (int i = 1; i <= exlRange->Columns->Count; i++)
			{
				dt->Columns->Add("колонка №" + i);
			}

			for (int i = 15; i <= exlRange->Rows->Count; i++)
			{
				DataRow^ row = dt->NewRow();
				for (int j = 1; j <= exlRange->Columns->Count; j++)
				{
					row[j - 1] = safe_cast<Excel::Range^>(exlRange->Cells[i, j])->Value2;
				}
				dt->Rows->Add(row);
			}
			dataGridView1->DataSource = dt;
			exlBook->Close(false, Type::Missing, Type::Missing);
			exlApp->Quit();
		}
	}

	double profit = 0; //нужно для расчетов
	double profitability = 0; //сохраняем в блокнот
	private: System::Void buttonRent_Click(System::Object^ sender, System::EventArgs^ e) {
		if (CheckName("ОТЧЕТ О ФИНАНСОВЫХ РЕЗУЛЬТАТАХ*",2,1))
		{
				System::Double::TryParse((safe_cast<Excel::Range^>(exlRange->Cells[23, 32])->Value2)->ToString(), profit);

				double costprice;
				System::Double::TryParse((safe_cast<Excel::Range^>(exlRange->Cells[19, 32])->Value2)->ToString(), costprice);
				

				profitability = (profit / costprice) * 100;
				labelProfitability->Text = "Уровень рентабельности = " + RoundDouble(profitability, 2);
		}
			exlBook->Close(false, Type::Missing, Type::Missing);
			exlApp->Quit();
	}

	double fixedAssets = 0;//нужно для расчетов
	double profitRate = 0; //сохраняем в блокнот
	private: System::Void buttonProfitRate_Click(System::Object^ sender, System::EventArgs^ e) {
		if (profit == 0)
		{
			MessageBox::Show("Вначале посчитайте рентабельность", "Ошибка последовательности",
				MessageBoxButtons::OK, MessageBoxIcon::Warning);
		}
		else
		{
			if (CheckName("БУХГАЛТЕРСКИЙ БАЛАНС*", 2, 1)) {
				System::Double::TryParse((safe_cast<Excel::Range^>(exlRange->Cells[24, 6])->Value2)->ToString(), fixedAssets);
				double currentAssets = 0;
				System::Double::TryParse((safe_cast<Excel::Range^>(exlRange->Cells[37, 6])->Value2)->ToString(), currentAssets);

				double workingСapital = fixedAssets + currentAssets;
				profitRate = (profit / workingСapital) * 100;
				labelProfitRate->Text = "Норма прибыли = " + RoundDouble(profitRate, 2);
			}
		}
	}

	
double сapitalLaborRatio = 0;//сохраняем в блокнот
private: System::Void buttonCapitalLaborRatio_Click(System::Object^ sender, System::EventArgs^ e) {
		if (fixedAssets == 0)
		{
			MessageBox::Show("Вначале посчитайте норму прибыли", "Ошибка последовательности",
				MessageBoxButtons::OK, MessageBoxIcon::Warning);
		}
		else
		{
			if (CheckName("ОТЧЕТ О ЧИСЛЕННОСТИ И ЗАРАБОТНОЙ ПЛАТЕ", 1, 1))
			{
				int meanCountWorker = 0;
				System::Int32::TryParse((safe_cast<Excel::Range^>(exlRange->Cells[18, 3])->Value2)->ToString(), meanCountWorker);
				сapitalLaborRatio = fixedAssets / meanCountWorker;
				labelCapitalLaborRatio->Text = "Фондовооруженность = " + RoundDouble(сapitalLaborRatio,2);
			}
		}
	}

double Ko = 0; //сохраняем в блокнот
private: System::Void buttonKo_Click(System::Object^ sender, System::EventArgs^ e) {
	if (fixedAssets == 0)
	{
		MessageBox::Show("Вначале посчитайте норму прибыли", "Ошибка последовательности",
			MessageBoxButtons::OK, MessageBoxIcon::Warning);
	}
	else
	{
		if (CheckName("ОТЧЕТ ОБ ОТРАСЛЕВЫХ ПОКАЗАТЕЛЯХ ДЕЯТЕЛЬНОСТИ", 1, 1))
		{
			double cashProceeds = 0;
			System::Double::TryParse((safe_cast<Excel::Range^>(exlRange->Cells[118, 4])->Value2)->ToString(), cashProceeds);

			Ko = cashProceeds / fixedAssets;
			labelKo->Text = "Коэффициент оборачиваемости = " + RoundDouble(Ko, 2);
		}
	}
}


	   //реализовать сохранение//
	private: System::Void buttonSave_Click(System::Object^ sender, System::EventArgs^ e) {
		saveFileDialog1->Filter = "(текстовый файл(*.txt)| *.txt";
		saveFileDialog1->RestoreDirectory = true;
		try
		{
			if (saveFileDialog1->ShowDialog() == System::Windows::Forms::DialogResult::OK)
			{
				String^ outputText = "Уровень рентабельности = " + RoundDouble(profitability, 2) + "\n";
				outputText += "Норма прибыли = " + RoundDouble(profitRate, 2) + "\n";
				outputText += "Фондовооруженность = " + RoundDouble(сapitalLaborRatio, 2) + "\n";
				outputText += "Коэффициент оборачиваемости = " + RoundDouble(Ko, 2);
				IO::File::WriteAllText(saveFileDialog1->FileName, outputText);
				MessageBox::Show(this, "Вы сохранили", saveFileDialog1->FileName,
					MessageBoxButtons::OK, MessageBoxIcon::Information);
			}
		}
		catch (Exception^ exception)
		{
			MessageBox::Show(this, "не удалось открыть файл", exception->Message,
				MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
};
}
