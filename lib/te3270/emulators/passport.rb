module TE3270
  module Emulators

    class Passport

      attr_writer :session_file, :visible, :window_state, :max_wait_time

      def initialize
        if jruby?
          require 'jruby-win32ole'
          require 'java'
          include_class 'java.awt.Dimension'
          include_class 'java.awt.Rectangle'
          include_class 'java.awt.Robot'
          include_class 'java.awt.Toolkit'
          include_class 'java.awt.event.InputEvent'
          include_class 'java.awt.image.BufferedImage'
          include_class 'javax.imageio.ImageIO'
        else
          require 'win32ole'
          require 'win32/screenshot'
        end
      end

      def connect
        start_passport_system

        yield self if block_given?
        raise 'The session file must be set in a block when calling connect with the Passport emulator.' if @session_file.nil?
        open_session
        @screen = session.Screen
        @area = screen.SelectAll
      end

      def disconnect
        session.Close
      end

      def get_string(row, column, length)
        screen.GetString(row, column, length)
      end

      def put_string(str, row, column)
        screen.PutString(str, row, column)
        quiet_period
      end

      def send_keys(keys)
        screen.SendKeys(keys)
        quiet_period
      end

      def wait_for_string(str, row, column)
        wait_for do
          screen.WaitForString(str, row, column)
        end
      end

      def wait_for_host(seconds)
        wait_for(seconds) do
          screen.WaitHostQuiet
        end
      end

      def wait_until_cursor_at(row, column)
        wait_for do
          screen.WaitForCursor(row, column)
        end
      end

      def screenshot(filename)
        File.delete(filename) if File.exists?(filename)
        session.Visible = true unless visible

        if jruby?
          toolkit = Toolkit::getDefaultToolkit()
          screen_size = toolkit.getScreenSize()
          rect = Rectangle.new(screen_size)
          robot = Robot.new
          image = robot.createScreenCapture(rect)
          f = java::io::File.new(filename)
          ImageIO::write(image, "png", f)
        else
          hwnd = session.WindowHandle
          Win32::Screenshot::Take.of(:window, hwnd: hwnd).write(filename)
        end

        session.Visible = false unless visible
      end

      def text
        area.Value
      end

      private

      attr_reader :system, :sessions, :session, :screen, :area

      WINDOW_STATES = {
          minimized: 0,
          normal: 1,
          maximized: 2
      }

      def wait_for(seconds = system.TimeoutValue / 1000)
        wait_collection = yield
        wait_collection.Wait(seconds * 1000)
      end

      def quiet_period
        wait_for_host(max_wait_time)
      end

      def max_wait_time
        @max_wait_time ||= 1
      end

      def window_state
        @window_state.nil? ? 1 : WINDOW_STATES[@window_state]
      end

      def visible
        @visible.nil? ? true : @visible
      end

      def open_session
        @sessions = system.Sessions
        @session = sessions.Open @session_file
        @session.WindowState = window_state
        @session.Visible = visible
      end

      def start_passport_system
        begin
          @system = WIN32OLE.new('PASSPORT.System')
        rescue Exception => e
          $stderr.puts e
        end
      end

      def jruby?
        RUBY_PLATFORM == 'java'
      end
    end
  end
end
