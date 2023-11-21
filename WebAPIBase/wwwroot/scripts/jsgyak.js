window.onload = function () {
    for (var sor = 0; sor < 10; sor++) {

        var újdiv = document.createElement("div");
        újdiv.classList.add("sor");
        document.getElementById("pascal").appendChild(újdiv);

        for (var oszlop = 0; oszlop < sor; oszlop++) {

            var újelemdiv = document.createElement("div");
            újelemdiv.classList.add("elem")
            újelemdiv.innerText = faktoriális(sor) / (faktoriális(oszlop) * faktoriális(sor - oszlop));
            újdiv.appendChild(újelemdiv);
        }
    }
}

var faktoriális = function (n) {
    let er = 1;
    for (let i = 2; i <= n; i++) {
        er = er * i;
    }
    return er;
}