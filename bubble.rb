# 1.冒泡排序
def bubble_sort(arr)
  1.upto(arr.length-1) do |i|
    (arr.length-i).times do |j|
      if arr[j]>arr[j+1]
        arr[j],arr[j+1] = arr[j+1],arr[j]
      end
    end
  end
    arr
end

array = [2.3,1.3,15.02,25.02,45,85.14,56.1,35.2,4.2,15.4]
puts bubble_sort(array)
