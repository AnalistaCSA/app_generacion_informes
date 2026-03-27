import { useState } from "react";
import usuarios from "../data/usuarios.json";
import "../styles/LoginPages.css";
import csa from "../img/csa.png"

function Login({ onLogin} ){
    const [username,setUsername] = useState("");
    const [password,setPassword] = useState("");

    const handleLogin = () =>{
        const user = usuarios.find(
            (u) => u.username === username && u.password === password
        );

        if (user){
            onLogin(user);
        } else {
            alert("Credenciales Incorrectas");
        }
    };

    return(
        <div class="contenedor">
            <img class="logo-csa" src={csa} alt="logo"/>
            <h1>Generador de informes</h1>
            <h2>Inicio de Sesión</h2>
            <br/>
            <br/>
            <input type="text" placeholder="Usuario" value={username} onChange={(e) => setUsername(e.target.value)}/>
            <input type="password" placeholder="Contraseña" value={password} onChange={(e) => setPassword(e.target.value)}/>
            <br/>
            <br/>
            <button class="boton-sesion" onClick={handleLogin}>Ingresar</button>
        </div>
    );
};

export default Login;


