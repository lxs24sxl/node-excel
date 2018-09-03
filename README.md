# 基于node-xlsx的操作excel文件的方法

### 构造函数初始化

```javascript
  let Excel = require('./excel/index.js')

  let excel = new Excel('./assets/test.xlsx', '订单详情')
```

#### add 向表格末尾添加一个数据(后面可拓展)
* @param {Array} arr 参数 如下

```javascript
excel.add ( [
    '腾讯新闻',
    201808,
    9997,
    'Tencent',
    '北京',
    '腾讯新闻-视频闪屏-动态展示全屏点击-外链(2017)',
    'mobile',
    'DISPLAY',
    '电视剧',
    'LXS2018DSP003HH',
  ],
);
```

### remove 删除指定表格中的某一项
* @param {Object} target 参数 如下 

```javascript
let target = { col_name: '订单ID', col_data: 9994};
excel.remove(target)
```


### find 查找表格中的某一项
* @param {Object} target 参数 如下 

```javascript
let target = { col_name: '订单ID', col_data: 9994};
excel.find(target)
```

### update 更新某行数据的某些数据
* @param {Object} target 参数 如下 
* @param {Array} newVal 参数 如下 

```javascript
let target = { col_name: '订单ID', col_data: 9994};

let newVal = [{ col_name: '市场', col_data: '广州'}, {col_name: '频道', col_data: '电影'}]

excel.update(target, newVal)
```