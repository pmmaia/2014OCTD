namespace :generate do
  desc "Generating data"
  task :data => [:students, :update]
  task :students => :environment do 
    100.times{ |i| Student.create!(fname: "Composite SS_#{i+1}", lname: "SS_#{i+1}-0.5", marks: rand(1500), grade: "Composite") }
  end
  task :update => :environment do
    students = Student.all
    pb_ttlc = 1000
    wetpb_stlc = 5
    students.each do |s|
      p = rand(80)/10.0
      r = ((s.marks > pb_ttlc) || p > wetpb_stlc) ? "California Hazardous" : "Non Hazardous" 
      s.update_attributes(percentage: p, remark: r)
    end
  end
end
