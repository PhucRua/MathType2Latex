# mt2mml.rb
require 'mathtype_to_mathml'

# Gem kỳ vọng FILE PATH, không phải bytes
path = ARGV[0]
abort "usage: ruby mt2mml.rb <oleObject*.bin>" unless path && File.exist?(path)

# Trả về MathML (string)
puts MathTypeToMathML::Converter.new(path).convert
