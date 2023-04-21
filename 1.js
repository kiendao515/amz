class Bug {
    constructor (name, phrase, power) {
        this.name = name
        this.phrase = phrase
        this.power = power
        this.species = "bug"
    }
    hide = () => console.log("You can't catch me now!")
    sayPhrase = () => console.log(this.phrase)
    attack = () => console.log(`I'm attacking with a power of ${this.power}!`)
}

class Robot {
    constructor (name, phrase, power) {
        this.name = name
        this.phrase = phrase
        this.power = power
        this.species = "robot"
    }
    transform = () => console.log("Optimus prime!")
    sayPhrase = () => console.log(this.phrase)
    attack = () => console.log(`I'm attacking with a power of ${this.power}!`)
}

const bug1 = new Bug("Buggy", "Your debugger doesn't work with me!", 10)
const Robot1 = new Robot("Tito", "I can cook, swim and dance!", 15)

console.log(bug1.power) //output: 10
Robot1.attack() // output: "I'm attacking with a power of 15!"