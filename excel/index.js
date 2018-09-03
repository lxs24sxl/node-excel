let fs = require ('fs');
let xlsx = require ('node-xlsx');

function throwError (name) {
  throw new Error ('Missing parameter: ' + name);
}

function Excel (path = throwError ('path'), name) {
  this.path = path;
  let xlsxData = xlsx.parse (path);
  this.xlsxData = xlsxData;
  this.name = name || xlsxData[0].name;
}

/**
 * 向表格添加数据
 * @param {Array} arr 参数 
 */
Excel.prototype.add = function (arr) {
  let self = this;
  // 调用build方法，参数是一个数据，每个成员为一个对象
  // name 表示sheet表名， data表示表数据
  let curSheet = self.xlsxData.find (item => item.name === self.name);
  // 传入数据到最后面
  curSheet.data.push (arr);
  // 将xlsx转为buffer数据
  let buffer = xlsx.build (self.xlsxData);
  // 写入文件
  fs.writeFile (self.path, buffer, err => {
    // 错误
    if (err) return console.error (err);
    console.log ('数据写入成功!');
    console.log ('---------我是分割线----------');
    console.log ('读取写入的数据');
    let testData = xlsx.parse (self.path);
    console.log (testData);
  });
};

/**
 * 根据表头的数据，删除数据
 */
Excel.prototype.remove = function (arg) {
  let self = this;
  // 查找当前sheet
  let curSheet = self.xlsxData.find (item => item.name === self.name);
  // 过滤出当前表下标
  let col_index = curSheet.data[0].findIndex (
    item => item.replace (/\s+/g, '') === arg.col_name
  );
  // 查找当前数据下标
  let data_index = curSheet.data.findIndex (
    item => item[col_index] === arg.col_data
  );

  curSheet.data.splice (data_index, 1);

  let buffer = xlsx.build (self.xlsxData);

  fs.writeFile (self.path, buffer, err => {
    // 错误
    if (err) return console.error (err);
    console.log (`删除 ${arg.col_name}的${arg.col_data} 成功!`);
    console.log ('---------我是分割线----------');
    console.log ('读取写入的数据');
    let testData = xlsx.parse (self.path);
    console.log (testData);
  });
};

/**
 * 查找当前项数据
 */
Excel.prototype.find = function (arg) {
  let self = this;
  // 查找当前sheet
  let curSheet = self.xlsxData.find (item => item.name === self.name);
  if (!curSheet) {
    return {
      success: false,
      params: arg,
      msg: `找不到当前sheet名为 ${self.name} 的表`,
    };
  }
  // 过滤出当前表下标
  let col_index = curSheet.data[0].findIndex (
    item => item.replace (/\s+/g, '') === arg.col_name
  );
  if (col_index <= -1) {
    return {
      success: false,
      params: arg,
      msg: `找不到当前属性名为 ${arg.col_name} 的属性`,
    };
  }
  // 查找当前数据下标
  let data = curSheet.data.find (item => item[col_index] === arg.col_data);
  // 返回数据
  if (!data) {
    return {
      success: false,
      params: arg,
      msg: `未查找到 属性名: ${arg.col_name} 属性值: ${arg.col_data} 的数据`,
    };
  }
  return {
    success: true,
    data: data,
  };
};

/**
 * 
 * @param {Object} target {col_name: '属性名', col_data: '属性值'}
 * @param {Array} newVal [{col_name: '属性名', col_data: '属性值'},{col_name: '属性名', col_data: '属性值'}]
 */
Excel.prototype.update = function (
  target = throwError ('target'),
  newVal = throwError ('newVal')
) {
  let self = this;
  let oldVal = self.find (target);
  let curSheet = self.xlsxData.find (item => item.name === self.name);
  newVal.map (nItem => {
    let oIndex = curSheet.data[0].findIndex (
      item => item.replace (/\s+/g, '') === nItem.col_name
    );
    oldVal.data[oIndex] = nItem.col_data;
    return nItem;
  });
  
  let buffer = xlsx.build( self.xlsxData );

  fs.writeFile(self.path, buffer, err => {
    // 错误
    if (err) return console.error (err);
    let msg = newVal.reduce((result, item) => {
      result.push(`${item.col_name}:${item.col_data}`)
      return result
    }, []).join(',')
    console.log (`更改 ${target.col_name}的${target.col_data} 为 ${msg} 成功!`);
    console.log ('---------我是分割线----------');
    console.log ('读取写入的数据');
    let testData = xlsx.parse (self.path);
    console.log (testData);
  })
};

module.exports = Excel;
