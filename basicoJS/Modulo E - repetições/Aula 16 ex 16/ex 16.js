var start = document.querySelector('#inicio')
var end = document.querySelector('#fim')
var step = document.querySelector('#passo')
var esc = document.querySelector('#escrever')

function passar(){
    var sta = Number(start.value)
    var ennd = Number(end.value)
    var ste = Number(step.value)
    esc.innerHTML = `<strong>Contando:</strong><br>`
    if (ste == 0 || ste == ''){
        alert('Está programado para que se o passo for igual a 0 ou sem numero, ele será considerado como 1')
        ste = 1
    }
    if (sta == '' || ennd == '' || ste == ''){
        alert('Preencha todos os campos!')
    } else if (sta > ennd){
        alert(`Espertinho, com o inicio maior que o fim não tem conta \u{1F92D}, mas eu tenho uma solução \u{1F447}`) 
        while(sta >= ennd){
            esc.innerHTML += `${sta}\u{1F449}`
            sta -= ste
    }
    } else{
        for ( sta ; sta <= ennd; sta += ste){ /* o primeiro sta é o inicio, algo do time 'Para sta(start)' logo em seguida vai para sta <= end e desce para completar a estrutura, depois de passar por toda a estrutura, volta para o ultimo comando, que é sta(start) + ele mesmo + ste(step)*/
            esc.innerHTML += `${sta}\u{1F449}`
        }
        esc.innerHTML += `\u{1F3C1}`
    }
}


