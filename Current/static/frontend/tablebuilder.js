


function generateTable(c=5, r=3){
    table = ""
    c = $("#colh").val() || c
    r = $("#rowh").val() || r
    t = $("#table")

    i = 0
    j = 0
    while(i < r){
        table += "<tr>"
        
        while(j < c){
            if(i == 0){
                if(j == 0){
                    table += `<th style="padding:3px 5px;"> <input id="h${j}" value="" type="text" placeholder="key" style="width:80px; color:black;"> </th>`

                }
                else{
                    table += `<th style="padding:3px 5px;"> <input id="h${j}" type="text" placeholder="header" style="width:80px; color:black;"> </th>`
                }
            }
            else{
                if(j == 0){
                    table += `<th style="padding:3px 5px;"> <input id="k${i}" value="" type="text" placeholder="name" style="width:80px; color:black;"> </th>`
                }else{
                    table += `<td style="padding:3px 5px;"> ${i},${j} </td>`  
                }
            }
            j++
        }
        table += "</tr>"
        i ++
        j = 0
    } 
    
    t.html(table)
    
}


$(document).ready(()=>{
    generateTable()
})


function create(c=5, r=3){
    t = $("#table");
    c = $("#colh").val() || c
    r = $("#rowh").val() || r


    console.log($("#k1").val(), $("#k2").val())
    
    jsonObject = {}

    i = 1
    keys = []
    while(i < r){
        
        keys.push($(`#k${i}`).val())
        i++
    }
    columns = []
    j = 0
    while(j < c){
        columns.push($(`#h${j}`).val())
        j++
    }

    jsonObject["Keys"] = keys
    
    jsonObject["Catagories"] = columns

    ajax(jsonObject)

}


function ajax(data){
    $.ajax({
        type: "POST",
        async: false,
        url: "/postmethod",
        contentType: "application/json",
        data: JSON.stringify(data),
        dataType: "json",
        success: function(response){
            console.table(response)
        },
        error : function(err){
            console.error(err);
        }
    })
}
