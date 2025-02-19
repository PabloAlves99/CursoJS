function carregar(){

var msg = document.querySelector('#msg')
var img = document.querySelector('#imagem')
var data = new Date()
var hora = data.getHours()
msg.innerHTML = `Agora são ${hora} hora`
    if (hora >= 0 && hora < 12){
        //Bom dia!
        img.src = 'manha.png'
        document.body.style.background = '#e2cd9f'
    } else if (hora >= 12 && hora <= 17){
        //Boa tarde!
        img.src ='tarde.png'
        document.body.style.background = '#b9846f'
    } else {
        //Boa noite!
        img.src ='noite.png'
        document.body.style.background = '#515154'
    }
}