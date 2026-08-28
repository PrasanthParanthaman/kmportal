document.getElementById('registrationForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent the default form submission

   
    // Get the user's name and email
    const id = document.getElementById('id').value;
    const password = document.getElementById('pass').value;

    // Store the name and email in localStorage
    // localStorage.setItem('userName', name);
    // localStorage.setItem('userEmail', email);


    // const role = document.getElementById('role').value;

    // if (role === 'admin') {
    //     window.location.href = "/HomePage"; // Redirect to index2 for admin
    // } else {
    //     alert('You do not have permission to register.'); // Alert for other roles
    // }

    if(id === '2112004' && password === 'Admin'){//prasanth
        window.location.href = '/HomePage';
    }
     else if(id === '2317709' && password === 'Admin'){//jagan
        window.location.href = '/HomePage';
     }
     else if(id === '285100' && password === 'Admin'){//monnisha
        window.location.href = '/HomePage';
     }
     else if(id === '2304018' && password === 'Admin'){//sindu
        window.location.href = '/HomePage';
     }
     else if(id === '319078' && password === 'Admin'){//Karthik
        window.location.href = '/HomePage';
     }
     else if(id === '448448' && password === 'Admin'){//uma
        window.location.href = '/HomePage';
     }
    else {
        alert('You do not have permission to register.'); // Alert for other roles
    }
});