let Excel = require ('./excel/index');

let excel = new Excel ('./assets/test.xlsx', '订单信息');
excel.add ([
    '腾讯新闻',
    201808,
    9980,
    'Tencent',
    '北京',
    '腾讯新闻-视频闪屏-动态展示全屏点击-外链(2017)',
    'mobile',
    'DISPLAY',
    '电视剧',
    'LXS2018DSP003HH',
  ],
);

// let target = { col_name: '订单ID', col_data: 9994};
// // excel.remove(target)
// // console.log (excel.find (target));

// let newVal = [{ col_name: '市场', col_data: '广州'}, {col_name: '频道', col_data: '电影'}]

// excel.update(target, newVal)
