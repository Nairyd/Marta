class Person:
    def __init__(self, lastname, firstname, confession):
        self.lastname = lastname
        self.firstname = firstname
        self.confession = confession

    def __str__(self):
        print(self.lastname, self.firstname, self.confession)

  #  def myfunction(self):
            #   do magic shit
