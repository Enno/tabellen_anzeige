class Array

  if RUBY_VERSION < '1.9'

    # Provides the cartesian product of two or more arrays.
    #
    #   a = []
    #   [1,2].product([4,5])
    #   a  #=> [[1, 4],[1, 5],[2, 4],[2, 5]]
    #
    # CREDIT: Thomas Hafner

    def product(*enums)
      enums.unshift self
      result = [[]]
      while [] != enums
        t, result = result, []
        b, *enums = enums
        t.each do |a|
          b.each do |n|
            result << a + [n]
          end
        end
      end
      result
    end

  end

  # Operator alias for cross-product.
  #
  #   a = [1,2] ** [4,5]
  #   a  #=> [[1, 4],[1, 5],[2, 4],[2, 5]]
  #
  # CREDIT: Trans

  alias_method :**, :product

end

