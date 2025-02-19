let num = document.querySelector('#number')
let lista = document.querySelector('#escrever')
let res = document.querySelector('#dados2')
let save = []

function isNumero (n){
    if(Number(n) >= 1 && Number(n) <= 100){
        return true
    } else {
        return false
    }
}

function inLista(n , l){
    if (l.indexOf(Number(n)) != -1 ){
        return true
    } else {
        return false
    }
}

function salvar(){
    if(isNumero(num.value) && !inLista(num.value, save)){
        save.push(Number(num.value))
        let item = document.createElement('option')
        item.text = `Valor ${num.value} adicionado`
        lista.appendChild(item)
        res.innerHTML = ''
    } else{
        alert('Valor inválido ou ja encontrado na lista.')
    }
    num.value = ''
    num.focus()
}

function escrever() {
    if (save.length == 0){
        alert('Adicione valores antes de finalizar')
    } else {
        let tot = save.length
        let maior = save[0]
        let menor = save[0]
        let soma = 0
        let media = 0
        for(let pos in save){
            soma += save[pos]
            if (save[pos] > maior)
                maior = save[pos]
            if (save[pos] < menor)
                menor = save[pos]
        }
        media = soma / tot
        res.innerHTML = ''
        res.innerHTML += `<p>Ao todo, temos ${tot} números cadastrados.</p>`
        res.innerHTML += `<p>O maior valor cadastrado foi ${maior}</p>`
        res.innerHTML += `<p> O menor valor cadastrado foi ${menor}</p>`
        res.innerHTML += `<p>Somando todos os valores temos ${soma}</p>`
        res.innerHTML += `<p>A média doso valores cadastrados é ${media}</p>`
    }
}
