$(()=>{
    //Post request to python server
    $("#upload-btn").click(()=>{
        var fData = new FormData($('#upload-files')[0])
        $.ajax({
            type: 'POST',
            url: '/sheetupload',
            data: fData,
            contentType: false,
            cache: false,
            processData: false,
            success: (data)=>{
                console.log('Success')
            },
            error: (e)=>{
                console.error(e)
            }
        })
    })
})

// file input and div the files selected are displayed in
const bf = $("#browseFiles")
const lf = $("#fileList")
// file icon
const fileIcon = '<span class="material-symbols-outlined" style="font-size:16px;color:white;">description</span>'

// function called when files are changed, displays the files selected in order of extension.
bf.change((e)=>{
    
    let filesSelected = e.target.files
    let list = []
    let returnHtml = ''

    //breaks down files into name and extension, and adds it to a running list
    for(var i = 0; i < filesSelected.length; i++){
        let name = filesSelected.item(i).name
        let spl = name.split(".")
        let ext = spl[spl.length-1]

        list.push([spl[0], ext])
    }

    // sorts the list in ascending alphabetical order (ods - xls - xlsx)
    list.sort((a, b) => a[1] > b[1] ? 1 : b[1] > a[1] ? -1 : 0)

    //builds the html that will be added to the site. Formatted like:
    // ICON .EXT &nbsp : NAME
    // &nbsp - non-breaking space (or tab)
    returnHtml += '<ul>'
    for(let li in list){
        returnHtml += `<li style="overflow:hidden;"> ${fileIcon} .${list[li][1]} &nbsp : ${list[li][0]} </li>`
    }
    returnHtml += '</ul>'

    lf.html(returnHtml)
})
