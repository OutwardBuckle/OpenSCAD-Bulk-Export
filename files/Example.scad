message = "X";
textsize = 20;
myfont = "Stencil";
height = 10;

linear_extrude(height){
    text(message, font=myfont,size=textsize);
}
