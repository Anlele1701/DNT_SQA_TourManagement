/* BEGIN EXTERNAL SOURCE */

        window.addEventListener("load", () => {
            const loader = document.querySelector(".loader");
            document.querySelector(".loader").classList.add("loader-hidden");
            document.querySelector(".loader").addEventListener("transitionend", () => {
                document.body.removeChild(document.querySelector(".loader"));
            });
        });


        function toggleDropdown() {
            var dropdown = document.getElementById("dropdown");
            if (dropdown.style.display === "block") {
                dropdown.style.display = "none";
            } else {
                dropdown.style.display = "block";
            }
        }
    
/* END EXTERNAL SOURCE */
/* BEGIN EXTERNAL SOURCE */

/* END EXTERNAL SOURCE */
