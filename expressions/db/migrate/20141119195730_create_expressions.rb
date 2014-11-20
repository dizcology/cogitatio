class CreateExpressions < ActiveRecord::Migration
  def change
    create_table :expressions do |t|

      t.timestamps
    end
  end
end
