var Trello = require("node-trello");
var _ = require('underscore');
var api = new Trello("9e1ba4f795ddca4d28c7ac2802725eef", "025fb59dc399193ff770962d34bf2ed1b7c8f8686ef27159d91a91d9bf926006");
var Workbook = require('xlsx-populate');

var cardNum = 4;

Workbook.fromFileAsync('/home/jacob/nodejs/trello/report.xlsx').then(
    workbook=> {
        // Modify the workbook.
        api.get('/1/boards/57a26cba986b92b41e5cd606/lists', function(err, data) {
        if(err) throw err;
        
            //console.log(data);
            _.each(data,function(element, idx, list){
                //console.log(element);
                
                api.get('/1/lists/' + element.id + '/cards', function(err, cards) {

                    _.each(cards,function(card, index, list){
                        //console.log(card);
                        if(card.idMembers && card.idMembers.length > 0 && card.idMembers.indexOf('53da30a45bc29f890fe836d0') != -1){
                            
                            //console.log(card.dateLastActivity);
                            var cardDate = new Date(card.dateLastActivity);
                            var currentTime = new Date();
                            currentTime.setDate(currentTime.getDate()-7);
                            
                            //console.log('Last two weeks ');
                            //console.log(currentTime);

                            if(cardDate.getDate() > currentTime.getDate()){
                                //console.log(card.dateLastActivity);
                                console.log(card.name);
                                //console.log(card.desc);
                                //console.log(cardNum);
                                
                                //console.log(workbook.sheet(0).cell('A4').value());
                                
                                if(element.name == 'Backlog' || element.name == 'Planning'　|| element.name == 'Released'){
                                    return;
                                }

                                //Label工时已报　
                                if(card.idLabels && card.idLabels.length > 0 && card.idLabels.indexOf('593d5b48ced82109ff0090fa') != -1){
                                    console.log('skip card' + card.name);
                                    return;
                                }

                                workbook.sheet(0).cell("B" + cardNum).value(card.name.replace('#','FOX - '));

                                if(element.name == 'In Progress'){
                                    workbook.sheet(0).cell("C" + cardNum).value('进行中');
                                }else{
                                    workbook.sheet(0).cell("C" + cardNum).value('完成');
                                }

                                //workbook.sheet(0).cell("D" + cardNum).value(card.desc);

                                cardNum++;
                            }
                        }
                    })

                    console.log('Before write')
                    workbook.toFileAsync("./out.xlsx");
                    
                });
        })
        
    });

    // Write to file.
    
});  


