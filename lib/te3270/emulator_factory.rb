require 'te3270/emulators/extra'
require 'te3270/emulators/quick3270'
require 'te3270/emulators/x3270'
require 'te3270/emulators/passport'

module TE3270
  #
  # Provides a mapping between a key used in the +emulator_for+ method
  # and the class that implements the access to the emulator.
  #
  module EmulatorFactory

    EMULATORS = {
        extra: TE3270::Emulators::Extra,
        quick3270: TE3270::Emulators::Quick3270,
        x3270: TE3270::Emulators::X3270,
        passport: TE3270::Emulators::Passport,
    }

    def self.emulator_for(platform)
      EMULATORS[platform]
    end

  end
end