const SENHA_PADRAO = "admin123";

function login() {
  const senha = document.getElementById("senha").value;

  if (senha === SENHA_PADRAO) {
    sessionStorage.setItem("logado", "true");
    window.location.href = "app.html";
  } else {
    document.getElementById("erro").innerText = "Senha incorreta";
  }
}