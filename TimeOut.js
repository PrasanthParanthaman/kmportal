// sessionTimeout.js
var inactivityTime = function () {
    var time;
    window.onload = resetTimer;
    // DOM Events
    document.onmousemove = resetTimer;
    document.onkeypress = resetTimer;

    function logout() {
        window.location.href = "/";
    }

    function resetTimer() {
        clearTimeout(time);
        time = setTimeout(logout, 60000);  // Redirect to index page after 1 minute
    }
};

window.onload = function() {
    inactivityTime();  // Start the timer on page load
};
