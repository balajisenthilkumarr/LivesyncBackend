export function calculateColumnMetrics(values, columnName = "") {
  const rowCount = values.length;
  // console.log(`[AGGREGATOR-v4] Single-pass metrics for "${columnName}" (${rowCount} rows).`);

  let sum = 0;
  let countNumeric = 0;
  let max = -Infinity;
  const valueCounts = {};
  let nonBlankCount = 0;

  const isDateColumn = columnName.toLowerCase().includes('date') || columnName.toLowerCase().includes('time');

  // Single Pass through the data
  for (let i = 0; i < rowCount; i++) {
    let val = values[i];
    
    // Handle Excel Object Wrapping
    if (val && typeof val === 'object' && val.hasOwnProperty('value')) val = val.value;
    
    if (val === null || val === undefined || String(val).trim() === "") continue;
    
    nonBlankCount++;
    const stringVal = String(val).trim();
    
    // 1. Categorical Accumulation (Always track for fallback/top-result)
    valueCounts[stringVal] = (valueCounts[stringVal] || 0) + 1;

    // 2. Numeric Accumulation (Skip if it's explicitly a Date column)
    if (!isDateColumn) {
      let num = typeof val === 'number' ? val : NaN;
      if (isNaN(num) && typeof val === 'string') {
        const clean = val.replace(/[$,]/g, '').trim();
        num = parseFloat(clean);
      }
      
      if (!isNaN(num)) {
        sum += num;
        countNumeric++;
        if (num > max) max = num;
      }
    }
  }

  // Decision Logic: Is it primarily numeric?
  const isNumeric = countNumeric > 0 && countNumeric >= (nonBlankCount * 0.5) && !isDateColumn;

  if (isNumeric) {
    return {
      type: 'numeric',
      count: countNumeric,
      sum: parseFloat(sum.toFixed(2)),
      avg: parseFloat((sum / countNumeric).toFixed(2)),
      max: max === -Infinity ? 0 : max
    };
  }

  // Fallback to Categorical
  const sortedCategories = Object.entries(valueCounts)
    .sort(([, a], [, b]) => b - a)
    .slice(0, 3)
    .map(([name, count]) => ({ name, count }));

  return {
    type: 'categorical',
    count: nonBlankCount,
    unique: Object.keys(valueCounts).length,
    top1: sortedCategories[0]?.name || 'None',
    top1_count: sortedCategories[0]?.count || 0,
    categories: sortedCategories
  };
}

export async function aggregateMergedMetrics(sheetMetrics, columnConfig) {
  const results = {};

  for (const config of columnConfig) {
    const colName = typeof config === 'string' ? config : config.name;
    const type = typeof config === 'string' ? 'add' : (config.type || 'add');
    
    const colMetrics = sheetMetrics.map(m => m[colName]).filter(m => m !== null && m !== undefined);
    
    if (colMetrics.length === 0) continue;

    const totalCount = colMetrics.reduce((a, b) => a + b.count, 0);
    
    // Applying Math Power: Add vs Minus
    const totalSum = colMetrics.reduce((a, b) => {
      return type === 'minus' ? a - b.sum : a + b.sum;
    }, 0);

    const totalAvg = totalCount > 0 ? totalSum / totalCount : 0;
    const totalMax = Math.max(...colMetrics.map(m => m.max));

    results[colName] = {
      count: totalCount,
      sum: parseFloat(totalSum.toFixed(2)),
      avg: parseFloat(totalAvg.toFixed(2)),
      max: type === 'minus' ? -totalMax : totalMax, // Reflect direction in Max too
      sheetCount: colMetrics.length,
      operation: type
    };
  }

  return results;
}
