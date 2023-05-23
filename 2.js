function findSumCombinations(arr, targetSum) {
  const combinations = [];

  function findSubsets(startIndex, currentCombination, currentSum) {
    if (currentSum === targetSum) {
      combinations.push(currentCombination.slice());
    }

    for (let i = startIndex; i < arr.length; i++) {
      currentCombination.push(arr[i]);
      currentSum += arr[i];
      findSubsets(i + 1, currentCombination, currentSum);
      currentCombination.pop();
      currentSum -= arr[i];
    }
  }

  findSubsets(0, [], 0);
  return combinations;
}

// Sử dụng hàm findSumCombinations với dữ liệu mẫu
const numbers = [-1, -1, 1, -1, -2, -3, 1, -1];
const targetSum = -3;

const result = findSumCombinations(numbers, targetSum);
console.log("Các cách tổng các số bằng", targetSum, "là:");
result.forEach(combination => {
  console.log(combination);
});
