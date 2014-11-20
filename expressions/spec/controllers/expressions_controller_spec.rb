require_relative "../rails_helper.rb"

RSpec.describe ExpressionsController, :type => :controller do
  before :each do
    @ec = ExpressionsController.new
  end
  describe "#swaps" do
    context "makes a list of swaps" do
    
      it "makes a list of three elements if the term has coefficient and variable" do
        expect(@ec.swaps("2x")).to eq(["2x","2*x","x*2"])
      end
    
      it "puts a variable without coefficient into an array" do
        expect(@ec.swaps("x")).to eq(["x"])
      end
    
      it "puts a number into an array" do
        expect(@ec.swaps("5")).to eq(["5"])
      end
    end
  
  end
  
  describe "#reps" do
    it "makes a list of representations of a term with sign" do
      term = "2x"
      expect(@ec.reps(term)).to eq(["+2x", "+2*x", "+x*2"])
    end
    
    
  end
  
  
  describe "#redundant" do
    it "extends an array by removing leading +" do
      a = ["+2x","3x","+3*y"]
      expect(@ec.redundant(a)).to eq((a+["2x","3*y"]).sort)
    end
  end

  describe "#assemble" do
    it "makes a list of all combinations of elements in the value arrays of the input hash" do
      h = {}
      h[1] = ["a", "b"]
      h[2] = ["x"]
      h[3] = ["c", "d"]
      
      expect(@ec.assemble(h)).to eq(%w(axc axd bxc bxd).sort)

    end
    
    it "returns the values of the hash if there is only one key" do
      h = {1 => [1,2,3]}
      
      expect(@ec.assemble(h)).to eq(h[1])
    end
    
    it "returns the glued pairs of the values in both orders if there are two keys" do
      h = {1=>["a","b","c"], 2=>["x","y"]}
      
      expect(@ec.assemble(h)).to eq(%w(ax ay bx by cx cy xa xb xc ya yb yc).sort)
    end
  end

end