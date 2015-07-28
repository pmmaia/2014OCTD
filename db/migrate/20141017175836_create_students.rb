class CreateStudents < ActiveRecord::Migration
  def change
    create_table :students do |t|
      t.string :fname
      t.string :lname
      t.integer :marks
      t.float :percentage
      t.string :grade
      t.string :remark

      t.timestamps
    end
  end
end
