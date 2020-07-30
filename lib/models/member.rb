# frozen_string_literal: true

class Member
  attr_accessor :name, # 姓名
                :gender, # 性別
                :job, # 職業
                :birth, # 出生
                :id_number, # 身份證字號
                :education, # 教育程度
                :address, # 住址
                :phone, # 電話
                :marriage, # 結婚日期與地點
                :marriage_witness, # 證婚人
                :spouse, # 配偶
                :father, # 父
                :mother, # 母
                :children, # 子女
                :baptize_date, # 洗禮日期
                :baptize_priest_and_church, # 施洗牧師/教會
                :first_communion_date, # 陪餐日期
                :first_communion_priest_and_church, # 接納牧師/教會
                :service_note, # 教會服事
                :training_note, # 訓練紀錄
                :visit_note, # 探訪註記
                :remark, # 備註
                :file_path
  def self.csv_header
    %w[姓名 性別 職業 出生 身份證字號 教育程度 住址 電話 結婚日期與地點 證婚人 配偶 父 母 子女 洗禮日期 施洗牧師/教會 陪餐日期 接納牧師/教會 教會服事 訓練紀錄 探訪註記 備註]
  end

  def initialize; end

  def load_data_from_docx(docx)
    self.name = docx.paragraphs.first.to_s.strip
    table = docx.tables.first
    raise ArgumentError, 'Cant load data from docx which have no table' if table.nil?

    self.gender = fetch_from_docx_table(table, 0, 1)
    self.job = fetch_from_docx_table(table, 0, 3)
    self.birth = fetch_from_docx_table(table, 0, 5)
    self.id_number = fetch_from_docx_table(table, 1, 1)
    self.education = fetch_from_docx_table(table, 1, 3)
    self.address = fetch_from_docx_table(table, 2, 1)
    self.phone = fetch_from_docx_table(table, 3, 1)
    self.marriage = fetch_from_docx_table(table, 4, 1)
    self.marriage_witness = fetch_from_docx_table(table, 4, 3)
    self.spouse = fetch_from_docx_table(table, 5, 1)
    self.father = fetch_from_docx_table(table, 5, 3)
    self.mother = fetch_from_docx_table(table, 5, 5)
    self.children = fetch_from_docx_table(table, 6, 1)
    self.baptize_date = fetch_from_docx_table(table, 8, 1)
    self.baptize_priest_and_church = fetch_from_docx_table(table, 8, 3)
    self.first_communion_date = fetch_from_docx_table(table, 9, 1)
    self.first_communion_priest_and_church = fetch_from_docx_table(table, 9, 3)
    self.service_note = fetch_from_docx_table(table, 10, 1)
    self.training_note = fetch_from_docx_table(table, 11, 1)
    self.visit_note = fetch_from_docx_table(table, 12, 1)
    self.remark = fetch_from_docx_table(table, 13, 1)
  end

  def to_csv
    [
      name,
      gender,
      job,
      birth,
      id_number,
      education,
      address,
      phone,
      marriage,
      marriage_witness,
      spouse,
      father,
      mother,
      children,
      baptize_date,
      baptize_priest_and_church,
      first_communion_date,
      first_communion_priest_and_church,
      service_note,
      training_note,
      visit_note,
      remark
    ]
  end

  private

  def fetch_from_docx_table(table, row, cell)
    table.rows.fetch(row) { |_| { cells: [] } }
         .cells.fetch(cell) { |_| { text: '' } }
         .text.strip
  end
end
