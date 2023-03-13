let pri = [5,8,4,7,6,3]
pri.push(1)
pri.sort()
console.log(pri)
console.log(`O vetor tem ${pri.length} posições`)
console.log(`O primeiro valor do vetor é ${pri[0]}`)

/*for (let teste = 0; teste < pri.length ; teste++){
    console.log(`A posição ${teste} tem o valor ${pri[teste]}`)
}*/

for(let teste in pri){
    console.log(`A posição ${teste} tem o valor ${pri[teste]}`)
}