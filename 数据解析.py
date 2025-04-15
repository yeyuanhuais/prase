var x = [];
var y1 = [];
var y2 = [];
var y3 = [];
var y4 = [];
for (var i = 0; i < 2000; i++) {
    var xName = '内部X轴.X' + (i);
    x.push({
        fullName: xName,
        value: 0
    });
    var y1Name = '曲线1.Y' + (2000 + i);
    y1.push({
        fullName: y1Name,
        value: 0
    });
    var y2Name = '曲线2.Y' + (4000 + i);
    y2.push({
        fullName: y2Name,
        value: 0
    });
    var y3Name ='曲线3.Y' + (6000 + i);
    y3.push({
        fullName: y3Name,
        value: 0
    });
    var y4Name = '曲线4.Y' + (8000 + i);
    y4.push({
        fullName: y4Name,
        value: 0
    });
}
Variable.BatchSetByFullNames(x, true);
Variable.BatchSetByFullNames(y1, true);
Variable.BatchSetByFullNames(y2, true);
Variable.BatchSetByFullNames(y3, true);
Variable.BatchSetByFullNames(y4, true);
$内部.x轴最大值 = 0;