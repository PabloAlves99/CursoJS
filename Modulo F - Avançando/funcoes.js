function chamada(){
    var ret = document.querySelector('#escrever')
    var nume = document.querySelector('#num')
    var nume = Number(nume.value)
    ret.innerHTML = verificar(nume)
}

function verificar(n){
    if (n % 2 == 0){
        return 'O numero é <strong>PAR</strong>'
    } else{
        return 'O numero é <strong>IMPAR</strong>'
    }
}

