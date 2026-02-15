export function calculateColumnMetrics(values) {
  // Strict numeric detection: Filter out anything that isn't a valid number
  const numbers = values
    .map(val => {
      if (typeof val === 'number') return val;
      if (typeof val === 'string') {
        const clean = val.replace(/[$,]/g, ''); // Basic currency/separator stripping
        const n = parseFloat(clean);
        return isNaN(n) ? null : n;
      }
      return null;
    })
    .filter(val => val !== null);

  // If the majority of values aren't numeric, we shouldn't return metrics for this column
  if (numbers.length === 0 || numbers.length < (values.length * 0.1)) {
    return null; // Return null to indicate this is likely a text/label column
  }

  const sum = numbers.reduce((a, b) => a + b, 0);
  const avg = sum / numbers.length;
  const max = Math.max(...numbers);

  return {
    count: numbers.length,
    sum: parseFloat(sum.toFixed(2)),
    avg: parseFloat(avg.toFixed(2)),
    max: parseFloat(max.toFixed(2))
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
