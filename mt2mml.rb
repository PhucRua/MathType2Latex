# Chuyển nhị phân MathType (MTEF) → MathML
# gem cần: gem install mathtype_to_mathml
require 'mathtype_to_mathml'
bin = File.binread(ARGV[0])
puts MathTypeToMathML::Converter.new(bin).convert
