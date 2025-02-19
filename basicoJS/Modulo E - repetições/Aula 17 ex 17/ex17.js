function gerar(){
    let num = document.querySelector('#nut')
    let tab = document.querySelector('#tabuada')
    
    if (num.value.length == ''){
        alert('Por favor digite um numero')
    } else{
        let n = Number(num.value)   
        let c = 1
        tab.innerHTML = ''
        for (n; c <= 10; c++){
            let item = document.createElement('option')
            item.text = `${n} x ${c} = ${n*c}`
            item.value = `tab${c}`
            tab.appendChild(item)
        }
    } 
}